from datetime import datetime
import os
import re
from struct import unpack
from uuid import UUID

MAXREGSID = 0xfffffffa
NOSTREAM = 0xffffffff

UNALLOCATED = 0x00
STORAGE_OBJECT = 0x01
STREAM_OBJECT = 0x02
ROOT_STORAGE_OBJECT = 0x05

RED = 0x00
BLACK = 0x01

EPOCH_AS_FILETIME = 116444736000000000L  # January 1, 1970 as MS file time
HUNDREDS_OF_NANOSECONDS = 10000000

from reader import VERSION_3, ENDOFCHAIN

class DirectoryEntry(object):
    def __init__(self, id, reader, position):
        self.entry_id = id
        self.reader = reader

        self.reader.id.seek(position)

        (self.name, directory_entry_name_length, self.object_type, \
            self.color_flag, self.left_sibling_id, self.right_sibling_id, \
            self.child_id, self.clsid, self.state_bits, self.creation_time, \
            self.modified_time, self.starting_sector_location, \
            self.stream_size) = unpack('<64sHBBLLL16sLQQLQ', self.reader.id.read(128))

        self.name = \
            self.name[:directory_entry_name_length].decode('utf-16').rstrip('\0')

        if re.search('[/\\:!]', self.name, re.UNICODE):
            raise Exception("The following characters are illegal and MUST "
                + "NOT be part of the name: '/', '\\', ':', '!'.");

        if self.object_type == UNALLOCATED:
            raise Exception("Can't create DirectoryEntry object for unallocated place")

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
        # I think most of files sneez on it...
        #if self.object_type == ROOT_STORAGE_OBJECT:
        #    if self.creation_time:
        #        raise Exception("For a root storage object, Creation Time "
        #            + "MUST be all zeroes, and the creation time is retrieved "
        #            + "or set on the compound file itself.")
        #    if self.modified_time:
        #        raise Exception("For a root storage object, Modified Time "
        #            + "MUST be all zeroes, and the modified time is retrieved "
        #            + "or set on the compound file itself.")

        if self.reader.major_version == VERSION_3 \
            and self.stream_size > 0x80000000:
            raise Exception("For a version 3 compound file 512-byte sector "
                + "size, this value of this field MUST be less than or equal "
                + "to 0x80000000")

        self.seek(0)

    def __del__(self):
        # HardRef deleting
        self.reader = None

    @property
    def left_sibling(self):
        if self.left_sibling_id == NOSTREAM:
            return None
        return self.reader.get_entry_by_id(self.left_sibling_id)

    @property
    def right_sibling(self):
        if self.right_sibling_id == NOSTREAM:
            return None
        return self.reader.get_entry_by_id(self.right_sibling_id)

    @property
    def child(self):
        if self.child_id == NOSTREAM:
            return None
        return self.reader.get_entry_by_id(self.child_id)

    def __repr__(self):
        return u'<Cfb%s#%d %s>' % (self.__class__.__name__, \
            self.entry_id, repr(self.name))

    def read(self, size=None):
        if not size or size < 0:
            size = self.stream_size

        buffer = ""

        if self.object_type == ROOT_STORAGE_OBJECT \
            or self.stream_size >= self.reader.mini_stream_cutoff_size:
            while len(buffer) < size:
                if self._position >= self.stream_size:
                    break
                if self._sector_number == ENDOFCHAIN:
                    break

                to_read = size - len(buffer)
                to_end  = self.reader.sector_size - self._position_in_sector
                to_do   = min(to_read, to_end)
                buffer += self.reader.id.read(to_do)
                self._position += to_do

                if to_read >= to_end:
                    self._position_in_sector = 0
                    
                    self._sector_number = \
                        self.reader.get_next_fat_sector(self._sector_number)
                    sector_position = (self._sector_number + 1) << \
                        self.reader.sector_shift
                    self.reader.id.seek(sector_position)
                else:
                    self._position_in_sector += to_do
        else:
            while len(buffer) < size:
                if self._position >= self.stream_size:
                    break
                if self._sector_number == ENDOFCHAIN:
                    break

                to_read = size - len(buffer)
                to_end  = self.reader.mini_sector_size - self._position_in_sector
                to_do   = min(to_read, to_end)
                buffer += self.reader.root_entry.read(to_do)
                self._position += to_do

                if to_read >= to_end:
                    self._position_in_sector = 0

                    self._sector_number = \
                        self.reader.get_next_mini_fat_sector(self._sector_number)
                    sector_position = self._sector_number << \
                        self.reader.mini_sector_shift
                    self.reader.root_entry.seek(sector_position)
                else:
                    self._position_in_sector += to_do

        return buffer[:size]

    def tell(self):
        return self._position

    def seek(self, offset, whence=os.SEEK_SET):
        if whence == os.SEEK_CUR:
            offset += self.tell()
        elif whence == os.SEEK_END:
            offset = self.stream_size - offset

        self._position = offset

        self._sector_number = self.starting_sector_location
        if self.object_type == ROOT_STORAGE_OBJECT \
            or self.stream_size >= self.reader.mini_stream_cutoff_size:

            current_position = 0
            while self._sector_number != ENDOFCHAIN \
                and (current_position + 1) * self.reader.sector_size < offset:
                self._sector_number = self.reader.get_next_fat_sector(self._sector_number)
                current_position += 1

            self._position_in_sector = offset - current_position * self.reader.sector_size
            sector_position = (self._sector_number + 1) << self.reader.sector_shift
            sector_position += self._position_in_sector
            self.reader.id.seek(sector_position)
        else:
            current_position = 0
            while self._sector_number != ENDOFCHAIN \
                and (current_position + 1) * self.reader.mini_sector_size < offset:
                self._sector_number = self.reader.get_next_mini_fat_sector(self._sector_number)
                current_position += 1

            self._position_in_sector = offset - current_position * self.reader.mini_sector_size
            sector_position = self._sector_number << self.reader.mini_sector_shift
            sector_position += self._position_in_sector
            self.reader.root_entry.seek(sector_position)
            
    def _filetime2timestamp(self, filetime):
        return datetime.utcfromtimestamp((filetime - EPOCH_AS_FILETIME) / \
            HUNDREDS_OF_NANOSECONDS)
