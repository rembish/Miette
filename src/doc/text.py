from cfb.reader import Reader

class DocTextReader(Reader):
    def __init__(self, filename):
        super(DocTextReader, self).__init__(filename)
        self._word_document = None
        self._n_table = None

    @property
    def word_document(self):
        if not self._word_document:
            self._word_document = self.get_entry_by_name('WordDocument')
        return self._word_document

    @property
    def n_table(self):
        if not self._n_table:
            a_to_m = self.word_document.get_short(0x000a)
            n_table_name = '1Table' if (a_to_m & 0x0200) == 0x0200 else '0Table'
            self._n_table = self.get_entry_by_name(n_table_name)

        return self._n_table
