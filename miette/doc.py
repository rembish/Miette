import os
from struct import unpack

from miette.cfb import CfbReader


__all__ = ['DocReader']


class DocReader(CfbReader):
    def __init__(self, filename):
        """
            Microsoft Word Document Reader (no markup)

            Usage example:
            >>> from miette.doc import DocReader
            >>> doc = DocReader('document.doc')
            >>> print doc.read()
            >>> print doc.word_document.get_short(0x000a)
        """
        super(DocReader, self).__init__(filename)
        self._word_document = None
        self._n_table = None

        self.cp = []
        self.length = None
        self._start_of_pcd = None

        self._read_clx()

        self._position = 0
        self.seek(0)

    @property
    def word_document(self):
        """
            Tunneling for WordDocument stream
        """
        if not self._word_document:
            self._word_document = self.get_entry_by_name('WordDocument')
            w_ident = self._word_document.get_short(0x0000)
            if w_ident != 0xa5ec:
                raise Exception('wIdent is an unsigned integer that specifies '
                                + 'that this is a Word Binary File. This value MUST be 0xA5EC.')
        return self._word_document

    @property
    def n_table(self):
        """
            Tunneling for 0Table/1Table stream
        """
        if not self._n_table:
            a_to_m = self.word_document.get_short(0x000a)
            n_table_name = '1Table' if (a_to_m & 0x0200) == 0x0200 else '0Table'
            self._n_table = self.get_entry_by_name(n_table_name)

        return self._n_table

    def read(self, size=None):
        """
            Read at most size bytes from .doc file (less if the read hits EOF
            before obtaining size bytes). If the size argument is negative or
            omitted, read all data until EOF is reached. The bytes are returned
            as a string object. An empty string is returned when EOF is
            encountered immediately.
        """

        if size is None:
            size = self.length - self.tell()

        dataBuffer = ""

        for i in range(len(self.cp) - 1):
            if self.cp[i + 1] < self.tell():
                continue

            self.n_table.seek(self._start_of_pcd + i * 8 + 2)
            fc = unpack('<L', self.n_table.read(4))[0]

            length = self.cp[i + 1] - self.cp[i]
            fc_f_compressed = (fc & 0x40000000) == 0x40000000
            fc_fc = fc & 0x3fffffff

            fc_fc += (self.tell() - self.cp[i]) * (2 if not fc_f_compressed else 1)
            length -= (self.tell() - self.cp[i]) * (1 if not fc_f_compressed else 2)
            if length > (size - len(dataBuffer)):
                length = size - len(dataBuffer)

            if fc_f_compressed:
                fc_fc /= 2
            else:
                length *= 2

            self.word_document.seek(fc_fc)
            part = self.word_document.read(length)
            if not fc_f_compressed:
                part = part.decode('utf-16')
            dataBuffer += part
            self._position += len(part)

            if len(dataBuffer) >= size:
                break

        return dataBuffer[:size].encode('utf-8')

    def tell(self):
        """
            Return the .doc's current position, like file's tell().
        """
        return self._position

    def seek(self, offset, whence=os.SEEK_SET):
        """
            Set the .doc's current position, like file's seek(). The whence
            argument is optional and defaults to os.SEEK_SET or 0 (absolute
            stream positioning); other values are os.SEEK_CUR or 1 (seek
            relative to the current position) and os.SEEK_END or 2 (seek
            relative to the .doc's end). There is no return value.
        """
        if whence == os.SEEK_SET:
            self._position = offset
        elif whence == os.SEEK_CUR:
            self._position += offset
        elif whence == os.SEEK_END:
            self._position = self.length - offset

        if self._position < 0:
            self._position = 0
        elif self._position > self.length:
            self._position = self.length

    def _read_clx(self):
        """
            Internal CLX work
        """
        self.word_document.seek(0x004c)
        (ccp_text, ccp_ftn, ccp_hdd, ccp_mcr, ccp_atn, ccp_edn, ccp_txbx,
         ccp_hdr_txbx) = unpack('<LLLLLLLL', self.word_document.read(32))

        last_cp = ccp_ftn + ccp_hdd + ccp_mcr + ccp_atn + ccp_edn + ccp_txbx \
                  + ccp_hdr_txbx
        last_cp += (0 if not last_cp else 1) + ccp_text

        pos = fc_clx = self.word_document.get_long(0x01a2)
        lcb_clx = self.word_document.get_long(0x01a6)
        clxt = self.n_table.get_byte(fc_clx)

        if clxt == 0x01:
            pos += 1
            cb_grpprl = self.n_table.get_short(pos)

            pos += 2
            if cb_grpprl > 0x3fa2:
                raise Exception('cbGrpprl MUST be less than or equal to 0x3FA2.')

            pos += cb_grpprl
            clxt = self.n_table.get_byte(pos)

        if clxt == 0x02:
            pos += 1
            lcb = self.n_table.get_long(pos)

            fc_plc_pcd = pos + 4
            if lcb != lcb_clx - 5:
                raise Exception('Wrong size of PlcPcd structure')
        else:
            raise Exception('PlcPcd.clxt MUST be 0x02')

        self.n_table.seek(fc_plc_pcd)
        for i in range(0, lcb_clx - (fc_plc_pcd - fc_clx), 4):
            self.cp.append(unpack('<L', self.n_table.read(4))[0])
            if self.cp[-1] == last_cp:
                self._start_of_pcd = self.n_table.tell()
                self.length = self.cp[-1] - self.cp[0]

                return

        raise Exception('Last found CP MUST be equal to lastCP')
