# Miette

Miette is a "small sweet thing" in French.

In another way, Miette is a light-weight, low-memory-usage library for reading
Microsoft Office documents — starting with Word Binary Files (`.doc`).

Requires Python 3.12+ and the [cfb](https://github.com/rembish/cfb) library.

## Usage

```python
from miette import DocReader

doc = DocReader("document.doc")
print(doc.read())
```

## Development

```bash
make install      # set up virtualenv + install dev dependencies
make format       # run black
make lint         # run ruff
make typecheck    # run mypy
make test         # run pytest with coverage
make pre-commit   # install pre-commit hooks
make clean        # remove build artifacts and caches
```

> **Note:** `cfb` is not yet on PyPI. The Makefile installs it from `../cfb`.
> For CI, it is installed from GitHub.

## License

BSD 2-Clause — see [LICENSE](LICENSE).
