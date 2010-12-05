import os
from os.path import basename
from struct import unpack

CLSID_NULL = '\0' * 16

VERSION_3 = 0x0003
VERSION_4 = 0x0004

ENDOFCHAIN = 0xfffffffe
FREESECT = 0xffffffff

from entry import Entry

class CfbReader(object):
    def __init__(self, filename):
        '''
            Compound File Binary File Format Reader

            Usage example:
            >>> cfb = Reader('document.doc')
            >>> word_document = cfb.root_entry.child.left_sibling
            >>> print word_document.read()

            >>> one_table = cfb.get_entry_by_name('1Table')
            >>> one_table.seek(100)
            >>> print one_table.read(16)
            >>> print one_table.tell()
        '''
        self.filename = filename
        self.id = open(self.filename, 'rb')

        header_signature = unpack('>Q', self.id.read(8))[0]
        if header_signature != 0xd0cf11e0a1b11ae1:
            raise Exception("Header signature MUST be set to the value 0xD0, "
                + "0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1.")

        header_clsid = self.id.read(16)
        if header_clsid != CLSID_NULL:
            raise Exception("Header CLSID reserved and unused class ID that "
                + "MUST be set to all zeroes (CLSID_NULL)")

        (minor_version, self.major_version, byte_order, self.sector_shift, \
            self.mini_sector_shift) = unpack('<HHHHH', self.id.read(10))

        if self.major_version not in (VERSION_3, VERSION_4):
            # Minor Version SHOULD (but MUST NOT) be set to 0x003E if the
            # major version field is either 0x0003 or 0x0004.
            raise Exception("Major Version MUST be set to either 0x0003 "
                + "(version 3) or 0x0004 (version 4).")

        if byte_order != 0xfffe:
            raise Exception("Byte Order MUST be set to 0xFFFE.")

        if (self.major_version == VERSION_3 and self.sector_shift != 0x0009) or \
            (self.major_version == VERSION_4 and self.sector_shift != 0x000c):
            raise Exception("Sector Shift MUST be set to 0x0009, or 0x000c, "
                + "depending on the Major Version field.")

        if self.mini_sector_shift != 0x0006:
            raise Exception("Mini Sector Shift MUST be set to 0x0006.")

        reserded = self.id.read(6)
        if reserded != '\0' * 6:
            raise Exception("Reserved field MUST be set to all zeroes.")

        (self.number_of_directory_sectors, self.number_of_fat_sectors, \
            self.first_directory_sector_location, self.transaction_signature_number, \
            self.mini_stream_cutoff_size, self.first_mini_fat_sector_location, \
            self.number_of_mini_fat_sectors, self.first_difat_sector_location, \
            self.number_of_difat_sectors) = unpack('<LLLLLLLLL', self.id.read(36))

        if self.major_version == VERSION_3 and self.number_of_directory_sectors != 0:
            raise Exception("If Major Version is 3, then the Number of "
                + "Directory Sectors MUST be zero. This field is not supported "
                + "for version 3 compound files.")

        if self.mini_stream_cutoff_size != 0x00001000:
            raise Exception("Mini Stream Cutoff Size MUST be set to 0x00001000.")

        self.sector_size = 1 << self.sector_shift
        self.mini_sector_size = 1 << self.mini_sector_shift

        self._root_entry = None
        self._directory = {}

    def __del__(self):
        self.id.close()

    def __repr__(self):
        return u'<%s %s@%d>' % (self.__class__.__name__, \
            basename(self.filename), self.tell())

    @property
    def root_entry(self):
        '''
            Tunneling to root entry storage
        '''
        if not self._root_entry:
            sector_number = self.first_directory_sector_location
            sector_position = (sector_number + 1) << self.sector_shift
            self._root_entry = Entry(0, self, sector_position)
            self._directory[0] = self._root_entry

        return self._root_entry

    def read(self, size=None):
        return self.id.read(size if size else -1)

    def seek(self, offset, whence=os.SEEK_SET):
        return self.id.seek(offset, whence)

    def tell(self):
        return self.id.tell()

    def get_entry_by_id(self, entry_id):
        '''
            Get directory entry object (stream or storage) by it's id
        '''
        if entry_id in self._directory:
            return self._directory[entry_id]

        sector_number = self.first_directory_sector_location
        current_entry = 0
        while sector_number != ENDOFCHAIN and \
            (current_entry + 1) * (self.sector_size / 128) <= entry_id:
            sector_number = self._get_next_fat_sector(sector_number)
            current_entry += 1

        sector_position = (sector_number + 1) << self.sector_shift
        sector_position += (entry_id - current_entry * (self.sector_size / 128)) * 128
        self._directory[entry_id] = Entry(entry_id, self, sector_position)
        return self._directory[entry_id]

    def get_entry_by_name(self, name):
        '''
            Get directory entry object (stream or storage) by it's name.
            Implements Microsoft Red-Black tree search algorithm. See
            [MS-CFB].doc @ 2.6.4 (pages 27-28).
        '''
        if self.root_entry.name == name:
            return self.root_entry
        current = self.root_entry.child

        while current:
            if len(current.name) < len(name):
                current = current.right_sibling
            elif len(current.name) > len(name):
                current = current.left_sibling
            elif cmp(current.name, name) < 0:
                current = current.right_sibling
            elif cmp(current.name, name) > 0:
                current = current.left_sibling
            else:
                return current

        return None

    def _get_next_fat_sector(self, current):
        '''
            Get next FAT sector block number
        '''
        difat_block = current / (self.sector_size / 4)
        if difat_block < 109:
            self.id.seek(76 + difat_block * 4)
        else:
            difat_block -= 109
            difat_sector = self.first_difat_sector_location

            while difat_block > (self.sector_size - 4) / 4:
                difat_position = (difat_sector + 1) << self.sector_shift
                self.id.seek(difat_position + self.sector_size - 4)
                difat_sector = unpack('<L', self.id.read(4))[0]
                difat_block -= (self.sector_size - 4) / 4

            difat_position = (difat_sector + 1) << self.sector_shift
            self.id.seek(difat_position + difat_block * 4)

        fat_sector = unpack('<L', self.id.read(4))[0]
        fat_sector_position = (fat_sector + 1) << self.sector_shift
        self.id.seek(fat_sector_position + (current % (self.sector_size / 4)) * 4)
        
        return unpack('<L', self.id.read(4))[0]

    def _get_next_mini_fat_sector(self, current):
        '''
            Get next mini FAT sector block number
        '''
        current_position = 0
        sector_number = self.first_mini_fat_sector_location
        
        while sector_number != ENDOFCHAIN and \
            (current_position + 1) * (self.sector_size / 4) <= current:
            sector_number = self._get_next_fat_sector(sector_number)
            current_position += 1

        if sector_number == ENDOFCHAIN:
            return ENDOFCHAIN

        sector_position = (sector_number + 1) << self.sector_shift
        sector_position += (current - current_position * (self.sector_size / 4)) * 4
        self.id.seek(sector_position)
        
        return unpack('<L', self.id.read(4))[0]
