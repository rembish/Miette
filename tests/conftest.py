"""Shared pytest fixtures for Miette tests."""

from __future__ import annotations

import warnings
from collections.abc import Iterator

import pytest

from miette import DocReader

_SIMPLE_DOC = "tests/data/simple.doc"


@pytest.fixture(autouse=True)
def suppress_warnings() -> Iterator[None]:
    """Suppress all Python warnings for the duration of each test."""
    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        yield


@pytest.fixture
def reader() -> Iterator[DocReader]:
    """Open simple.doc as a DocReader and close it after the test."""
    with DocReader(_SIMPLE_DOC) as r:
        yield r
