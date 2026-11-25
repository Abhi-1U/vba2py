"""Filesystem state and classic VBA numbered file channels."""

import csv
import glob
import os
import shutil
from decimal import Decimal
from pathlib import Path

from .errors import VBAError
from .values import EMPTY, MISSING, NULL, integer, invariant_number, number, string


def file_error(error):
    if isinstance(error, FileNotFoundError):
        return VBAError(53, str(error))
    if isinstance(error, PermissionError):
        return VBAError(70, str(error))
    if isinstance(error, FileExistsError):
        return VBAError(58, str(error))
    return VBAError(75, str(error))


class FileSystem:
    def __init__(self, directory=None, encoding="cp1252"):
        self.directory = Path(directory or Path.cwd()).resolve()
        self.encoding = encoding
        self.channels = {}
        self._matches = iter(())

    def path(self, name):
        raw = string(name)
        if not raw:
            raise VBAError(52)
        return (self.directory / raw.replace("\\", os.sep)).resolve()

    def free_file(self, RangeNumber=0):
        interval = integer(RangeNumber)
        if interval not in (0, 1):
            raise VBAError(5)
        for index in (range(1, 256) if interval == 0 else range(256, 512)):
            if index not in self.channels:
                return index
        raise VBAError(67)

    def open(self, name, mode, channel, access=None):
        index = integer(channel)
        if not 1 <= index <= 511:
            raise VBAError(52)
        if index in self.channels:
            raise VBAError(55)
        path = self.path(name)
        if mode == "random":
            raise VBAError(54, "Random-access VBA files are not supported")
        if (mode == "input" and access not in {None, "read"}
                or mode in {"output", "append"} and access not in {None, "write"}):
            raise VBAError(54, "The requested access is unsupported for the file mode")
        try:
            if mode == "binary":
                if access == "read":
                    handle = path.open("rb")
                else:
                    handle = path.open("r+b" if path.exists() else "w+b")
            else:
                handle = path.open({"input": "r", "output": "w", "append": "a"}[mode],
                                   encoding=self.encoding, newline="")
        except OSError as error:
            raise file_error(error) from None
        self.channels[index] = (handle, mode, path, [])

    def _channel(self, index, modes=None):
        try:
            result = self.channels[integer(index)]
        except KeyError:
            raise VBAError(52) from None
        if modes and result[1] not in modes:
            raise VBAError(54)
        return result

    def close(self, *channels):
        targets = tuple(integer(item) for item in channels) if channels else tuple(self.channels)
        for target in targets:
            record = self.channels.pop(target, None)
            if record is None:
                raise VBAError(52)
            try:
                record[0].close()
            except OSError as error:
                raise file_error(error) from None

    def print_line(self, channel, fields, write=False):
        handle = self._channel(channel, ("output", "append"))[0]

        def render(value):
            if value is NULL:
                return "#NULL#" if write else "Null"
            if value is EMPTY:
                return ""
            if write and isinstance(value, bool):
                return "#TRUE#" if value else "#FALSE#"
            if write and isinstance(value, str):
                return '"' + value.replace('"', '""') + '"'
            if isinstance(value, (int, float, Decimal)) and not isinstance(value, bool):
                if write:
                    return invariant_number(value)
                return (" " if value >= 0 else "") + string(value) + " "
            return string(value)

        try:
            if write:
                handle.write(",".join(render(value) for value, _ in fields) + "\r\n")
                return
            for value, separator in fields:
                handle.write(render(value))
                if separator == ",":
                    handle.write("\t")
            if not fields or not fields[-1][1]:
                handle.write("\r\n")
        except (OSError, UnicodeError) as error:
            raise VBAError(57, str(error)) from None

    def input(self, channel, count, line=False):
        handle, _, _, buffer = self._channel(channel, ("input",))
        try:
            if line:
                text = handle.readline()
                if not text:
                    raise VBAError(62)
                return [text.rstrip("\r\n")]
            while len(buffer) < count:
                row = handle.readline()
                if not row:
                    raise VBAError(62)
                buffer.extend(next(csv.reader([row], skipinitialspace=True), []))
            tokens = buffer[:count]
            del buffer[:count]
        except (OSError, UnicodeError, csv.Error) as error:
            raise VBAError(57, str(error)) from None

        def typed(token):
            booleans = {"#true#": True, "#false#": False, "#null#": NULL, "": EMPTY}
            if token.casefold() in booleans:
                return booleans[token.casefold()]
            try:
                return number(token)
            except (ValueError, IndexError, VBAError):
                return token

        return [typed(token) for token in tokens]

    def eof(self, channel):
        file, mode, path, buffer = self._channel(channel, ("input", "binary"))
        return not buffer and file.tell() >= path.stat().st_size

    def loc(self, channel):
        file, mode, _, _ = self._channel(channel)
        offset = file.tell()
        return offset if mode == "binary" else offset // 128

    def lof(self, channel):
        handle, _, path, _ = self._channel(channel)
        handle.flush()
        return path.stat().st_size

    def seek(self, channel, Position=MISSING):
        file, _, _, buffer = self._channel(channel)
        if Position is MISSING:
            return file.tell() + 1
        offset = integer(Position)
        if offset < 1:
            raise VBAError(5)
        file.seek(offset - 1)
        buffer.clear()

    def input_chars(self, count, channel):
        file, _, _, _ = self._channel(channel, ("input", "binary"))
        amount = integer(count)
        if amount < 0:
            raise VBAError(5)
        data = file.read(amount)
        if len(data) < amount:
            raise VBAError(62)
        return data.decode(self.encoding) if isinstance(data, bytes) else data

    def dir(self, PathName=MISSING, Attributes=0):
        if PathName is not MISSING:
            path = self.path(PathName)
            match_directories = bool(integer(Attributes) & 16)
            self._matches = iter(sorted(
                item.name for match in glob.glob(str(path)) if
                (item := Path(match)).is_file() or match_directories and item.is_dir()
            ))
        return next(self._matches, "")

    def chdir(self, path):
        target = self.path(path)
        if not target.is_dir():
            raise VBAError(76)
        self.directory = target

    def mkdir(self, path):
        return self.path(path).mkdir()

    def rmdir(self, path):
        return self.path(path).rmdir()

    def kill(self, pattern):
        matches = glob.glob(str(self.path(pattern)))
        if not matches:
            raise VBAError(53)
        for filename in matches:
            Path(filename).unlink()

    def copy(self, source, destination):
        shutil.copyfile(self.path(source), self.path(destination))

    def rename(self, source, destination):
        origin, target = self.path(source), self.path(destination)
        if target.exists():
            raise VBAError(58)
        origin.rename(target)
