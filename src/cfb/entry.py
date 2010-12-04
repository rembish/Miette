import os
import re

from struct import unpack
from uuid import UUID
from datetime import datetime

MAXREGSID = 0xfffffffa
NOSTREAM = 0xffffffff

UNALLOCATED = 0x00
STORAGE_OBJECT = 0x01
STREAM_OBJECT = 0x02
ROOT_STORAGE_OBJECT = 0x05

RED = 0x00
BLACK = 0x01

from reader import VERSION_3, ENDOFCHAIN

class Entry(object):
    def __init__(self, id, reader, position):
        self.entry_id = id
        self._reader = reader

        self._reader.id.seek(position)

        (self.name, directory_entry_name_length, self.object_type, \
            self.color_flag, self.left_sibling_id, self.right_sibling_id, \
            self.child_id, self.clsid, self.state_bits, self.creation_time, \
            self.modified_time, self.starting_sector_location, \
            self.stream_size) = unpack('<64sHBBLLL16sLQQLQ', self._reader.id.read(128))

        self.name = \
            self.name[:directory_entry_name_length].decode('utf-16').rstrip('\0')

        if re.search('[/\\:!]', self.name, re.UNICODE):
            raise Exception("The following characters are illegal and MUST "
                + "NOT be part of the name: '/', '\\', ':', '!'.");

        if self.object_type == UNALLOCATED:
            raise Exception("Can't create Directory Entry object for unallocated place")

        if self.object_type not in (STORAGE_OBJECT, STREAM_OBJECT, ROOT_STORAGE_OBJECT):
            raise Exception("Object Type MUST be 0x00, 0x01, 0x02, or 0x05, "
                + "depending on the actual type of object.")

        if self.color_flag not in (RED, BLACK):
            raise Exception("Color Flag MUST be 0x00 (red) or 0x01 (black).")

        if MAXREGSID < self.left_sibling_id < NOSTREAM:
            raise Exception("Maximum regular stream ID is 0xFFFFFFFA (MAXREGSID)")
        if MAXREGSID < self.right_sibling_id < NOSTREAM:
            raise Exception("Maximum regular stream ID is 0xFFFFFFFA (MAXREGSID)")
        if MAXREGSID < self.child_id < NOSTREAM:
            raise Exception("Maximum regular stream ID is 0xFFFFFFFA (MAXREGSID)")

        self.clsid = UUID(bytes=self.clsid)
        if self.object_type == STREAM_OBJECT:
            if self.clsid.int != 0:
                raise Exception("In a stream object, CLSID field MUST be set "
                    + "to all zeroes.")
            if self.state_bits != 0:
                raise Exception("State bits field MUST be set to all zeroes "
                    + "for a stream object.")

        self.creation_time = self._filetime2timestamp(self.creation_time) \
            if self.creation_time else None
        self.modified_time = self._filetime2timestamp(self.modified_time) \
            if self.modified_time else None

        # Documentation error? [MS-CFB].pdf @ 26
        # I think most of files sneeze on it...
        #if self.object_type == ROOT_STORAGE_OBJECT:
        #    if self.creation_time:
        #        raise Exception("For a root storage object, Creation Time "
        #            + "MUST be all zeroes, and the creation time is retrieved "
        #            + "or set on the compound file itself.")
        #    if self.modified_time:
        #        raise Exception("For a root storage object, Modified Time "
        #            + "MUST be all zeroes, and the modified time is retrieved "
        #            + "or set on the compound file itself.")

        if self._reader.major_version == VERSION_3 \
            and self.stream_size > 0x80000000:
            raise Exception("For a version 3 compound file 512-byte sector "
                + "size, this value of this field MUST be less than or equal "
                + "to 0x80000000")

        self.is_mini_sector = self.object_type != ROOT_STORAGE_OBJECT \
            and self.stream_size < self._reader.mini_stream_cutoff_size
        self.sector_size = self._reader.sector_size if not self.is_mini_sector \
            else self._reader.mini_sector_size
        self.sector_shift = self._reader.sector_shift if not self.is_mini_sector \
            else self._reader.mini_sector_shift

        self.seek(0)

    def __del__(self):
        # HardRef deleting
        self._reader = None

    def __repr__(self):
        return u'<Cfb%s#%d %s>' % (self.__class__.__name__, \
            self.entry_id, repr(self.name))

    @property
    def left_sibling(self):
        '''
            Tunneling to left sibling
        '''
        if self.left_sibling_id == NOSTREAM:
            return None
        return self._reader.get_entry_by_id(self.left_sibling_id)

    @property
    def right_sibling(self):
        '''
            Tunneling to right sibling
        '''
        if self.right_sibling_id == NOSTREAM:
            return None
        return self._reader.get_entry_by_id(self.right_sibling_id)

    @property
    def child(self):
        '''
            Tunneling to child (root entry only)
        '''
        if self.child_id == NOSTREAM:
            return None
        return self._reader.get_entry_by_id(self.child_id)

    @property
    def reader(self):
        '''
            Link to stream reader: file object for standard streams
            and root entry storage for mini streams
        '''
        return self._reader.id if not self.is_mini_sector \
            else self._reader.root_entry

    def read(self, size=None):
        '''
            Read at most size bytes from the stream (less if the read hits EOS
            before obtaining size bytes). If the size argument is negative or
            omitted, read all data until EOS is reached. The bytes are returned
            as a string object. An empty string is returned when EOS is
            encountered immediately.
        '''
        if not size or size < 0:
            size = self.stream_size - self.tell()

        buffer = ""
        while len(buffer) < size:
            if self._position >= self.stream_size:
                break
            if self._sector_number == ENDOFCHAIN:
                break

            to_read = size - len(buffer)
            to_end  = self.sector_size - self._position_in_sector
            to_do   = min(to_read, to_end)
            buffer += self.reader.read(to_do)
            self._position += to_do

            if to_read >= to_end:
                self._position_in_sector = 0

                self._sector_number = \
                    self._get_next_sector(self._sector_number)
                sector_position = (self._sector_number + \
                    int(not self.is_mini_sector)) << self.sector_shift
                self.reader.seek(sector_position)
            else:
                self._position_in_sector += to_do

        return buffer[:size]

    def tell(self):
        '''
            Return the stream's current position, like file's tell().
        '''
        return self._position

    def seek(self, offset, whence=os.SEEK_SET):
        '''
            Set the stream's current position, like file's seek(). The whence
            argument is optional and defaults to os.SEEK_SET or 0 (absolute
            stream positioning); other values are os.SEEK_CUR or 1 (seek
            relative to the current position) and os.SEEK_END or 2 (seek
            relative to the stream's end). There is no return value.
        '''
        if whence == os.SEEK_CUR:
            offset += self.tell()
        elif whence == os.SEEK_END:
            offset = self.stream_size - offset

        self._position = offset
        self._sector_number = self.starting_sector_location
        current_position = 0

        while self._sector_number != ENDOFCHAIN \
            and (current_position + 1) * self.sector_size < offset:
            self._sector_number = self._get_next_sector(self._sector_number)
            current_position += 1

        self._position_in_sector = offset - current_position * self.sector_size
        sector_position = (self._sector_number + int(not self.is_mini_sector)) \
            << self.sector_shift
        sector_position += self._position_in_sector
        
        self.reader.seek(sector_position)

    def get_byte(self, start):
        self.seek(start)
        return unpack('<B', self.read(1))[0]

    def get_short(self, start):
        self.seek(start)
        return unpack('<H', self.read(2))[0]

    def get_long(self, start):
        self.seek(start)
        return unpack('<L', self.read(4))[0]

    def _get_next_sector(self, current):
        '''
            Wrapper to the reader source get_next_fat_sector/get_next_mini_fat_sector
            methods. Transparent to all stream sizes.
        '''
        return self._reader._get_next_fat_sector(current) if not self.is_mini_sector \
            else self._reader._get_next_mini_fat_sector(current)
            
    def _filetime2timestamp(self, filetime):
        '''
            Convert Microsoft OLE time to datetime object
            116444736000000000L is January 1, 1970
        '''
        return datetime.utcfromtimestamp((filetime - 116444736000000000L) / 10000000)
