"""Tests for miette.exceptions."""

import pytest

from miette.exceptions import MietteError, MietteFormatError


def test_miette_error_is_exception() -> None:
    assert issubclass(MietteError, Exception)


def test_format_error_is_miette_error() -> None:
    assert issubclass(MietteFormatError, MietteError)


def test_raise_miette_error() -> None:
    with pytest.raises(MietteError):
        raise MietteError("base error")


def test_raise_miette_format_error() -> None:
    with pytest.raises(MietteFormatError):
        raise MietteFormatError("format error")


def test_format_error_caught_as_miette_error() -> None:
    with pytest.raises(MietteError):
        raise MietteFormatError("caught as parent")
