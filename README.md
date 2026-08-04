# pyArea — documentation source

Hebrew user guide for the [pyArea](https://github.com/blliav/pyArea) Revit extension.

**Live site:** https://blliav.github.io/pyArea/

This branch (`docs`) is an orphan branch — it shares no history with `master`/`dev`
and contains no extension code. Likewise, the code branches contain no documentation.

## Layout

```
mkdocs.yml     site config and navigation
docs/          the Hebrew markdown pages
```

## Editing

Edit the markdown under `docs/`, then commit and push to `docs`. The
`Deploy docs` GitHub Action builds the site and publishes it to the
`gh-pages` branch automatically — `gh-pages` is machine-managed, never edit it by hand.

## Local preview

Requires Python 3.13+ :

```
pip install mkdocs-material
mkdocs serve
```

Then open http://localhost:8000 . The preview reloads as you save.
