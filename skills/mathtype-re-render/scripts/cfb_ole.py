"""Small Compound File Binary (OLE2) reader for embedded MathType objects.

This intentionally implements only the read path needed to inspect the
Equation Native stream.  It uses no third-party packages, so it can run in
the constrained document-processing environment.
"""
from __future__ import annotations

import struct
from dataclasses import dataclass


FREESECT = 0xFFFFFFFF
ENDOFCHAIN = 0xFFFFFFFE
FATSECT = 0xFFFFFFFD
DISECT = 0xFFFFFFFC
NOSTREAM = 0xFFFFFFFF


@dataclass
class Entry:
    index: int
    name: str
    obj_type: int
    start: int
    size: int
    left: int
    right: int
    child: int


def u32(data: bytes, off: int) -> int:
    return struct.unpack_from("<I", data, off)[0]


def u64(data: bytes, off: int) -> int:
    return struct.unpack_from("<Q", data, off)[0]


class CompoundFile:
    def __init__(self, data: bytes):
        if data[:8] != bytes.fromhex("d0 cf 11 e0 a1 b1 1a e1"):
            raise ValueError("not an OLE2 compound file")
        self.data = data
        self.sector_size = 1 << struct.unpack_from("<H", data, 0x1E)[0]
        self.mini_sector_size = 1 << struct.unpack_from("<H", data, 0x20)[0]
        self.first_dir = u32(data, 0x30)
        self.first_minifat = u32(data, 0x3C)
        self.num_minifat = u32(data, 0x40)
        self.first_difat = u32(data, 0x44)
        self.num_difat = u32(data, 0x48)
        self.fat_sector_ids = self._read_difat()
        self.fat = self._read_fat()
        self.entries = self._read_directory()
        self.root = next(e for e in self.entries if e.obj_type == 5)
        self.minifat = self._read_minifat()
        self.mini_stream = self._read_regular(self.root.start, self.root.size)

    def sector(self, sid: int) -> bytes:
        off = (sid + 1) * self.sector_size
        return self.data[off : off + self.sector_size]

    def _read_difat(self) -> list[int]:
        out = []
        for i in range(109):
            sid = u32(self.data, 0x4C + i * 4)
            if sid != FREESECT:
                out.append(sid)
        sid = self.first_difat
        for _ in range(self.num_difat):
            if sid in (FREESECT, ENDOFCHAIN):
                break
            sec = self.sector(sid)
            for i in range(self.sector_size // 4 - 1):
                n = u32(sec, i * 4)
                if n != FREESECT:
                    out.append(n)
            sid = u32(sec, self.sector_size - 4)
        return out

    def _read_fat(self) -> list[int]:
        out: list[int] = []
        for sid in self.fat_sector_ids:
            sec = self.sector(sid)
            out.extend(u32(sec, i * 4) for i in range(self.sector_size // 4))
        return out

    def chain(self, start: int, table: list[int]) -> list[int]:
        out = []
        seen = set()
        sid = start
        while sid not in (FREESECT, ENDOFCHAIN, NOSTREAM):
            if sid in seen or sid >= len(table):
                raise ValueError(f"invalid sector chain at {sid:#x}")
            seen.add(sid)
            out.append(sid)
            sid = table[sid]
        return out

    def _read_regular(self, start: int, size: int) -> bytes:
        if size == 0 or start in (FREESECT, ENDOFCHAIN):
            return b""
        chunks = [self.sector(sid) for sid in self.chain(start, self.fat)]
        return b"".join(chunks)[:size]

    def _read_directory(self) -> list[Entry]:
        raw = self._read_regular(self.first_dir, 1 << 30)
        entries = []
        for idx, off in enumerate(range(0, len(raw) - 127, 128)):
            block = raw[off : off + 128]
            name_len = struct.unpack_from("<H", block, 0x40)[0]
            if name_len < 2 or name_len > 64:
                name = ""
            else:
                name = block[: name_len - 2].decode("utf-16le", "replace")
            entries.append(
                Entry(
                    idx,
                    name,
                    block[0x42],
                    u32(block, 0x74),
                    u64(block, 0x78),
                    u32(block, 0x44),
                    u32(block, 0x48),
                    u32(block, 0x4C),
                )
            )
        return entries

    def _read_minifat(self) -> list[int]:
        if self.first_minifat in (FREESECT, ENDOFCHAIN) or not self.num_minifat:
            return []
        raw = b"".join(self.sector(sid) for sid in self.chain(self.first_minifat, self.fat))
        return [u32(raw, i * 4) for i in range(len(raw) // 4)]

    def read_entry(self, entry: Entry) -> bytes:
        if entry.size < 4096:
            if not self.minifat:
                return b""
            chunks = []
            for sid in self.chain(entry.start, self.minifat):
                off = sid * self.mini_sector_size
                chunks.append(self.mini_stream[off : off + self.mini_sector_size])
            return b"".join(chunks)[: entry.size]
        return self._read_regular(entry.start, entry.size)

    def stream_entries(self) -> list[Entry]:
        return [e for e in self.entries if e.obj_type == 2]

    def find_stream(self, name: str) -> Entry | None:
        return next((e for e in self.stream_entries() if e.name == name), None)


def inspect(data: bytes) -> list[tuple[str, int, int]]:
    cfb = CompoundFile(data)
    return [(e.name, e.start, e.size) for e in cfb.stream_entries()]

