"""Miette exception hierarchy."""

__all__ = ["MietteError", "MietteFormatError"]


class MietteError(Exception):
    """Base exception for all Miette errors."""


class MietteFormatError(MietteError):
    """Raised when a document violates the Word Binary File Format spec."""
