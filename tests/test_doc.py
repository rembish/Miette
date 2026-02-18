"""Tests for miette.doc (DocReader)."""

from pathlib import Path
from unittest.mock import MagicMock, patch

import pytest
from cfb import CfbIO
from cfb.directory.entry import SEEK_CUR, SEEK_END

from miette import DocReader
from miette.exceptions import MietteFormatError

_SIMPLE_DOC = "tests/data/simple.doc"


def test_open_with_str_path(reader: DocReader) -> None:
    assert reader.length > 0


def test_open_with_pathlib_path() -> None:
    assert DocReader(Path(_SIMPLE_DOC)).length > 0


def test_initial_position(reader: DocReader) -> None:
    assert reader.tell() == 0


def test_read_returns_bytes(reader: DocReader) -> None:
    assert isinstance(reader.read(), bytes)


def test_read_size(reader: DocReader) -> None:
    assert len(reader.read(1)) <= 1


def test_read_zero(reader: DocReader) -> None:
    assert reader.read(0) == b""


def test_read_advances_position(reader: DocReader) -> None:
    reader.read(1)
    assert reader.tell() == 1


def test_read_full_content(reader: DocReader) -> None:
    assert reader.read() == b"One two three four five.\r"


def test_read_skips_past_blocks(reader: DocReader) -> None:
    reader.seek(reader.length // 2)
    assert isinstance(reader.read(), bytes)


def test_seek_set(reader: DocReader) -> None:
    reader.seek(0)
    assert reader.tell() == 0


def test_seek_end(reader: DocReader) -> None:
    reader.seek(0, SEEK_END)
    assert reader.tell() == reader.length


def test_seek_clamps_negative(reader: DocReader) -> None:
    reader.seek(-9999, SEEK_CUR)
    assert reader.tell() == 0


def test_seek_clamps_past_end(reader: DocReader) -> None:
    reader.seek(9999999)
    assert reader.tell() == reader.length


def test_context_manager() -> None:
    with DocReader(_SIMPLE_DOC) as reader:
        assert reader.length > 0
        assert not reader.io.closed
    assert reader.io.closed


def test_repr(reader: DocReader) -> None:
    assert repr(reader) == f'<DocReader "{_SIMPLE_DOC}">'


def test_n_table_name(reader: DocReader) -> None:
    assert reader.n_table.name == "1Table"


def test_bad_wident_raises(reader: DocReader) -> None:
    # Evict cached property so it will be re-evaluated.
    # Dunder methods must be patched at class level (Python special-method lookup
    # bypasses instance __dict__), so we temporarily override CfbIO.__getitem__.
    del reader.__dict__["word_document"]
    mock_entry = MagicMock()
    mock_entry.get_short.return_value = 0x0000
    with patch.object(CfbIO, "__getitem__", return_value=mock_entry), pytest.raises(
        MietteFormatError, match="wIdent"
    ):
        _ = reader.word_document
