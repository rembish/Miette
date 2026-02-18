"""Microsoft Word Document (.doc) reader — extracts plain text."""

from functools import cached_property
from pathlib import Path
from struct import unpack
from typing import Self

from cfb import CfbIO
from cfb.directory.entry import SEEK_CUR, SEEK_END, SEEK_SET, Entry

from miette.exceptions import MietteFormatError

__all__ = ["DocReader"]


class DocReader:
    """Read plain text from a Microsoft Word Binary File (.doc)."""

    def __init__(self, filename: str | Path) -> None:
        self.io = CfbIO(filename)

        self.cp: list[int] = []
        self.length: int = 0
        self._start_of_pcd: int = 0
        self._position: int = 0

        self._read_clx()
        self.seek(0)

    def __enter__(self) -> Self:
        return self

    def __exit__(self, *args: object) -> None:
        self.io.close()

    def __repr__(self) -> str:
        return f'<DocReader "{self.io.name}">'

    @cached_property
    def word_document(self) -> Entry:
        """WordDocument stream."""
        word_document = self.io["WordDocument"]
        w_ident = word_document.get_short(0x0000)
        if w_ident != 0xA5EC:
            raise MietteFormatError(
                "wIdent is an unsigned integer that specifies that this is a Word "
                "Binary File. This value MUST be 0xA5EC."
            )
        return word_document

    @cached_property
    def n_table(self) -> Entry:
        """0Table or 1Table stream."""
        a_to_m = self.word_document.get_short(0x000A)
        n_table_name = "1Table" if a_to_m & 0x0200 == 0x0200 else "0Table"
        return self.io[n_table_name]

    def read(self, size: int | None = None) -> bytes:
        """Read at most *size* bytes of UTF-8 encoded text from the document.

        If *size* is omitted or ``None``, read until EOF.
        """
        if size is None:
            size = self.length - self.tell()

        data_buffer = ""

        for i in range(len(self.cp) - 1):
            if self.cp[i + 1] < self.tell():
                continue

            self.n_table.seek(self._start_of_pcd + i * 8 + 2)
            fc = unpack("<L", self.n_table.read(4))[0]

            length = self.cp[i + 1] - self.cp[i]
            fc_f_compressed = (fc & 0x40000000) == 0x40000000
            fc_fc = fc & 0x3FFFFFFF

            fc_fc += (self.tell() - self.cp[i]) * (1 + (not fc_f_compressed))
            length -= (self.tell() - self.cp[i]) * (1 + fc_f_compressed)
            if length > (size - len(data_buffer)):
                length = size - len(data_buffer)

            if fc_f_compressed:
                fc_fc //= 2
                self.word_document.seek(fc_fc)
                part = self.word_document.read(length).decode("cp1252")
            else:
                self.word_document.seek(fc_fc)
                part = self.word_document.read(length * 2).decode("utf-16-le")

            data_buffer += part
            self._position += len(part)

            if len(data_buffer) >= size:
                break

        return data_buffer[:size].encode("utf-8")

    def tell(self) -> int:
        """Return the current position within the document."""
        return self._position

    def seek(self, offset: int, whence: int = SEEK_SET) -> None:
        """Set the current position within the document.

        *whence* follows the usual ``os`` constants: ``SEEK_SET``, ``SEEK_CUR``,
        ``SEEK_END``.
        """
        if whence == SEEK_SET:
            self._position = offset
        elif whence == SEEK_CUR:
            self._position += offset
        elif whence == SEEK_END:
            self._position = self.length - offset

        if self._position < 0:
            self._position = 0
        elif self._position > self.length:
            self._position = self.length

    def _read_clx(self) -> None:
        """Parse the CLX structure to populate the CP table and document length."""
        self.word_document.seek(0x004C)
        (
            ccp_text,
            ccp_ftn,
            ccp_hdd,
            ccp_mcr,
            ccp_atn,
            ccp_edn,
            ccp_txbx,
            ccp_hdr_txbx,
        ) = unpack("<LLLLLLLL", self.word_document.read(32))

        last_cp = (
            ccp_ftn + ccp_hdd + ccp_mcr + ccp_atn + ccp_edn + ccp_txbx + ccp_hdr_txbx
        )
        last_cp += (0 if not last_cp else 1) + ccp_text

        pos = fc_clx = self.word_document.get_long(0x01A2)
        lcb_clx = self.word_document.get_long(0x01A6)
        clxt = self.n_table.get_byte(fc_clx)

        if clxt == 0x01:
            pos += 1
            cb_grpprl = self.n_table.get_short(pos)
            pos += 2
            if cb_grpprl > 0x3FA2:
                raise MietteFormatError(
                    "cbGrpprl MUST be less than or equal to 0x3FA2."
                )
            pos += cb_grpprl
            clxt = self.n_table.get_byte(pos)

        if clxt != 0x02:
            raise MietteFormatError("PlcPcd.clxt MUST be 0x02")

        pos += 1
        lcb = self.n_table.get_long(pos)
        fc_plc_pcd = pos + 4
        if lcb != lcb_clx - 5:
            raise MietteFormatError("Wrong size of PlcPcd structure")

        self.n_table.seek(fc_plc_pcd)
        for _ in range(0, lcb_clx - (fc_plc_pcd - fc_clx), 4):
            self.cp.append(unpack("<L", self.n_table.read(4))[0])
            if self.cp[-1] == last_cp:
                self._start_of_pcd = self.n_table.tell()
                self.length = self.cp[-1] - self.cp[0]
                return

        raise MietteFormatError("Last found CP MUST be equal to lastCP")
