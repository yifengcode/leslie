# leslie

Convert a Microsoft Word `.docx` document to Markdown with Python.

## Install

```sh
python -m pip install -r requirements.txt
```

## Use

```sh
python word_to_markdown.py input.docx
```

The Markdown file is written next to the input with a `.md` extension. To
choose another output path:

```sh
python word_to_markdown.py input.docx --output output.md
```

The converter preserves headings, bold and italic text, lists, quotes, and
tables. It accepts `.docx` files; older `.doc` files should first be saved as
`.docx` in Word or another compatible editor.
