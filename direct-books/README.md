# Direct Books

Drop pre-parsed book JSON files here to view them in the viewer **without running the parser**.

They appear in a **separate dropdown** (the second selector in the topbar) — not mixed in with the parsed books from `output/`.

## Two ways to load a file

### 1. Register it in the dropdown

Add an entry to [`direct-books-index.json`](direct-books-index.json):

```json
[
  { "name": "My Book", "file": "my-book.json" },
  { "name": "Another",  "path": "direct-books/another.json" }
]
```

- `file` — filename inside `direct-books/` (shorthand)
- `path` — full path from the project root (if you want to keep the JSON elsewhere)
- `name` — label shown in the dropdown

The viewer reads this file on load, so refresh the page after editing.

### 2. Ad-hoc: the **Load JSON** button

Click **Load JSON** in the topbar and pick any `.json` file from disk. This works even without the index entry and without a local server.

## Expected JSON shape

Same shape the parser produces:

```json
{
  "toc":   [ ... ],
  "pages": [ { "label": "...", "originalLabels": [...], "html": "...", "sectionHeading": "..." }, ... ]
}
```
