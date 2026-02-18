# Changelog

## 2.0.1 (2026-02-18)

- Modernise test suite: autouse `suppress_warnings` with `catch_warnings()` + yield, `reader` fixture with proper teardown via context manager
- Add `test_read_full_content` and `test_n_table_name` tests
- Close superseded PR #2 (both its fixes were already in 2.0.0)

## 2.0.0 (2026-02-18)

- Dropped Python 2 and older Python 3 support; requires Python 3.12+
- Replaced `setup.py` with `pyproject.toml` (PEP 517/518)
- Added full type annotations and `py.typed` marker (PEP 561)
- Added `__enter__`/`__exit__` context manager support to `DocReader`
- Added `__repr__` to `DocReader`
- Switched `@cached` to `functools.cached_property` following cfb 0.9.x API
- Fixed Python 3 bytes decoding: compressed text decoded as `cp1252`,
  uncompressed as `utf-16-le`
- Fixed integer division bug: `fc_fc /= 2` → `fc_fc //= 2`
- Added pytest test suite with 90%+ coverage
- Added Black, Ruff, mypy (strict) tooling
- Added GitHub Actions CI and Dependabot configuration

## 1.5 (2013)

- Last 1.x release
