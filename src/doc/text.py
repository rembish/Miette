import os

from os.path import basename
from cfb.reader import Reader
from struct import unpack

class DocTextReader(Reader):
    def __init__(self, filename):
        '''
            Microsoft Word Document Reader (no markup)

            Usage example:
            >>> doc = DocTextReader('document.doc')
            >>> print doc.read()
            >>> print doc.word_document.get_short(0x000a)
        '''
        super(DocTextReader, self).__init__(filename)
        self._word_document = None
        self._n_table = None

    def __repr__(self):
        return u'<DocReader %s@%d>' % (basename(self.filename), self.id.tell())

    @property
    def word_document(self):
        '''
            Tunneling for WordDocument stream
        '''
        if not self._word_document:
            self._word_document = self.get_entry_by_name('WordDocument')
        return self._word_document

    @property
    def n_table(self):
        '''
            Tunneling for 0Table/1Table stream
        '''
        if not self._n_table:
            a_to_m = self.word_document.get_short(0x000a)
            n_table_name = '1Table' if (a_to_m & 0x0200) == 0x0200 else '0Table'
            self._n_table = self.get_entry_by_name(n_table_name)

        return self._n_table

    def read(self, size=None):
        '''
            read(size) prototype
        '''
        self.word_document.seek(0x004c)
        (ccp_text, ccp_ftn, ccp_hdd, ccp_mcr, ccp_atn, ccp_edn, ccp_txbx, \
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
            raise Exception('clxt MUST be 0x02')

        cp = []
        self.n_table.seek(fc_plc_pcd)
        for i in range(0, lcb_clx - (fc_plc_pcd - fc_clx), 4):
            cp.append(unpack('<L', self.n_table.read(4))[0])
            if cp[-1] == last_cp:
                break

        if not len(cp) or cp[-1] != last_cp:
            raise Exception('Last found CP MUST be equal to lastCP')

        parts = []
        for i in range(0, lcb_clx - (self.n_table.tell() - fc_clx), 8):
            (abc_fr2, fc, prm) = unpack('<HLH', self.n_table.read(8))
            fc_f_compressed = (fc & 40000000) == 40000000
            fc_fc = fc & 0x3fffffff

            j = i / 8
            length = cp[j + 1] - cp[j]
            if fc_f_compressed:
                fc_fc /= 2
            else:
                length *= 2

            parts.append((fc_fc, length, fc_f_compressed))

        buffer = ""
        for (start, length, compressed) in parts:
            self.word_document.seek(start)
            if not compressed:
                buffer += self.word_document.read(length).decode('utf-16')
            else:
                buffer += self.word_document.read(length)

        return buffer.encode('utf-8')

    def seek(self, offset, whence=os.SEEK_SET):
        raise NotImplementedError('seek')

    def tell(self):
        raise NotImplementedError('tell')
