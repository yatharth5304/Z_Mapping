"""Single-owner Excel checkpoints with bounded requests and immutable row chunks."""

from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime, timezone
import os
from pathlib import Path
import sys
import tempfile
import threading
from zipfile import ZIP_DEFLATED, ZipFile

from openpyxl.writer.excel import ExcelWriter


def save_workbook_atomic(workbook, destination):
    """Serialize beside the destination, then replace it only after a full save."""
    destination = Path(destination)
    destination.parent.mkdir(parents=True, exist_ok=True)
    if workbook.read_only:
        raise TypeError("Workbook is read-only")
    if workbook.write_only and not workbook.worksheets:
        workbook.create_sheet()
    temporary = None
    try:
        with tempfile.NamedTemporaryFile(
            prefix=".zuventus_", suffix=".xlsx", dir=destination.parent, delete=False,
        ) as handle:
            temporary = Path(handle.name)
        # Match openpyxl's save_workbook, but own the archive context explicitly.
        # Its normal save() can retain an open ZIP in a failed writer traceback,
        # preventing removal of the partial file on Windows.
        with ZipFile(temporary, "w", ZIP_DEFLATED, allowZip64=True) as archive:
            workbook.properties.modified = datetime.now(timezone.utc).replace(tzinfo=None)
            ExcelWriter(workbook, archive).save()
        os.replace(temporary, destination)
    finally:
        if temporary is not None:
            temporary.unlink(missing_ok=True)


@dataclass(frozen=True, slots=True, eq=False)
class _RowChunk:
    previous: _RowChunk | None
    rows: tuple


@dataclass(frozen=True, slots=True, eq=False)
class FrozenRows:
    """An immutable completed prefix; traversal never uses recursive calls."""

    tail: _RowChunk | None = None
    count: int = 0

    def __len__(self):
        return self.count

    def __iter__(self):
        chunks = []
        chunk = self.tail
        while chunk is not None:
            chunks.append(chunk.rows)
            chunk = chunk.previous
        for rows in reversed(chunks):
            yield from rows


class CheckpointSaveError(OSError):
    """The last completed prefix could not be committed during finalization."""


class AsyncWorkbookSaver:
    """One producer, one writer, and at most one latest pending checkpoint.

    append() freezes scalar cell values into tuples. Only the producer touches
    its short buffer; the writer receives immutable prefixes and owns all Excel
    objects. No producer lock is held during serialization, logging, or I/O.
    """

    def __init__(self, writer, interval=1000, on_progress=None, on_error=None):
        if isinstance(interval, bool) or not isinstance(interval, int) or interval < 0:
            raise ValueError("autosave interval must be a nonnegative integer (0 disables intermediate saves)")
        self.interval = interval
        self._writer = writer
        self._on_progress = on_progress
        self._on_error = on_error
        self._condition = threading.Condition()
        self._buffer = []
        self._tail = None
        self._count = 0
        self._pending = None
        self._stopping = False
        self._closed = False
        self._last_saved = -1
        self._last_error = None
        self._thread = threading.Thread(target=self._run, name="excel-autosave", daemon=False)
        self._thread.start()

    @property
    def worker_alive(self):
        return self._thread.is_alive()

    def _message(self, message, *, error=False):
        callback = self._on_error if error else self._on_progress
        if callback is None and not error:
            return
        try:
            if callback is not None:
                callback(message)
            else:
                print(message, file=sys.stderr, flush=True)
        except Exception:
            # A broken progress sink must not invalidate an otherwise saved file.
            try:
                print(message, file=sys.stderr, flush=True)
            except Exception:
                pass

    def append(self, row):
        if self._closed:
            raise RuntimeError("Cannot append after autosave finalization")
        self._buffer.append(tuple(row))
        self._count += 1
        if self.interval and self._count % self.interval == 0:
            self.checkpoint()

    def checkpoint(self):
        if self._closed:
            raise RuntimeError("Cannot request a checkpoint after finalization")
        if self._buffer:
            self._tail = _RowChunk(self._tail, tuple(self._buffer))
            self._buffer = []
        snapshot = FrozenRows(self._tail, self._count)
        with self._condition:
            # Replacing a pointer coalesces requests without copying the prefix.
            self._pending = snapshot
            self._condition.notify()

    def _run(self):
        try:
            while True:
                with self._condition:
                    self._condition.wait_for(lambda: self._pending is not None or self._stopping)
                    if self._pending is None:
                        return
                    snapshot = self._pending
                    self._pending = None
                if len(snapshot) == self._last_saved and self._last_error is None:
                    continue
                self._message(f"[autosave] Saving checkpoint at {len(snapshot)} processed rows...")
                try:
                    self._writer(snapshot)
                except Exception as exc:
                    self._last_error = exc
                    self._message(
                        f"[autosave] ERROR at {len(snapshot)} rows: {exc}. "
                        "The previous checkpoint is retained; a later checkpoint or final save will retry.",
                        error=True,
                    )
                else:
                    recovered = self._last_error is not None
                    self._last_saved = len(snapshot)
                    self._last_error = None
                    suffix = " (recovered after a save failure)" if recovered else ""
                    self._message(f"[autosave] Checkpoint saved: {len(snapshot)} processed rows{suffix}")
        finally:
            close = getattr(self._writer, "close", None)
            if close is not None:
                try:
                    close()
                except Exception as exc:
                    self._last_error = exc
                    self._message(f"[autosave] ERROR closing writer: {exc}", error=True)

    def finish(self):
        if not self._closed:
            self.checkpoint()
            with self._condition:
                self._closed = True
                self._stopping = True
                self._condition.notify()
        self._thread.join()
        if self._last_saved != self._count or self._last_error is not None:
            raise CheckpointSaveError(
                f"Final Excel checkpoint failed at {self._count} processed rows; "
                "the previous successfully saved file is retained"
            ) from self._last_error

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc_value, traceback):
        try:
            self.finish()
        except Exception as save_error:
            if exc_type is None:
                raise
            self._message(f"[autosave] Finalization also failed: {save_error}", error=True)
        return False
