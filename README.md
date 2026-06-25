# LaTeX to Word converter

Convert a LaTeX paper to Microsoft Word (`.docx`) with [pandoc](https://pandoc.org)
while keeping the mathematics, the figures and an IEEE-style bibliography, and
applying a Word style template.

The script handles the cases pandoc does not cover on its own:

- **IEEE bibliography.** Citations are numbered `[1], [2], ...` in order of
  appearance. The reference text is taken from the `.bbl` produced by
  `IEEEtran.bst`, and a DOI is appended to each entry (read from the `doi` field
  in the `.bib`).
- **Cross references.** `\ref` and `\eqref` are replaced by the numbers LaTeX
  already computed, read from the `.aux` file (so `Algorithm~\ref{...}` becomes
  `Algorithm 3`, `Eq.~\eqref{...}` becomes `Eq. (2)`, and so on).
- **Equations.** Displayed equations are centered and numbered with the number at
  the right margin, using a right-aligned tab (Word's native equation-numbering
  layout, not a table).
  Pseudocode (`algorithm` / `algorithmic`) gets a left-aligned header (heading,
  Input, Output) framed by a top and bottom rule, followed by the numbered steps in
  a two-column table (one row per line) so every line stays separate, while the
  inline math is preserved as real Word equations.
- **Figures.** Figures are found through `\graphicspath` and embedded as PNG.
- **Word styles.** When a template is supplied, pandoc uses it, the tables are
  centered, and the auto-generated paragraph styles (figure/table captions, author,
  date) are mapped onto the template styles.

## Requirements

- Python 3.8 or newer (standard library only).
- pandoc 2.11 or newer (3.x recommended).
- A compiled LaTeX project: the converter needs the `.aux` and `.bbl` files, so
  run LaTeX and BibTeX once before converting. The `.bbl` must come from a
  numeric, appearance-ordered style such as `IEEEtran`.

## Usage

```bash
# Convert the bundled example:
python latex_to_word.py main.tex -d ./example -t ./reference.docx
```

Options:

| Option | Meaning |
| --- | --- |
| `main.tex` | main `.tex` file (default: `main.tex`) |
| `-d`, `--paper-dir` | folder holding the LaTeX sources (default: current folder) |
| `-t`, `--template` | Word template; copied into the paper folder as `reference.docx` and used as the style template |
| `-o`, `--output` | output `.docx` path (default: `<main>.docx` inside the paper folder) |
| `--year-after-authors` | put the publication year in parentheses after the last author (author-year style, e.g. for HJS) instead of at the end of the entry |

By default the references keep the IEEE layout, with the year at the end of each
entry. Add `--year-after-authors` to get the author-year layout:

```
[1] J. Yin, Z. Zheng, and L. Cao (2012), "USpan: ...," in Proceedings ... pp. 660-668. DOI: ...
```

The same settings can be passed through the `L2W_MAIN`, `L2W_DIR`,
`L2W_TEMPLATE` and `L2W_YEAR_AFTER_AUTHORS` environment variables when you prefer
to set them once:

```bash
L2W_DIR=example L2W_MAIN=main.tex L2W_TEMPLATE=reference.docx python latex_to_word.py
```

### After the conversion

When it finishes, the converter prints a short review checklist: items found in
this run (references without a DOI, cross references it could not resolve, figures
that did not embed) followed by general reminders about the parts pandoc cannot
reproduce exactly (theorem numbering, equation-number alignment, pseudocode
layout, styles). Open the `.docx` in Word and go through the list.

### Google Colab

`Latex2WordHJS.ipynb` runs the same conversion on Google Colab. The first cell
installs a recent pandoc; the second cell has input boxes for the LaTeX folder,
the Word template and the main file; the third cell downloads the result.

## Repository layout

```
latex_to_word.py      the converter
Latex2WordHJS.ipynb   Google Colab notebook
reference.docx        Word style template (Title, Subtitle, Figure, Table, ...)
example/              minimal runnable sample paper
```

The `example/` folder is a small, self-contained paper (title, abstract, a
numbered equation, an algorithm, cross references and two citations). It already
ships with the `.aux` and `.bbl` files, so you can run the converter on it right
away. Point `-d` at your own LaTeX project to convert your paper instead.

## How the template works

pandoc maps document elements onto named paragraph styles. The supplied
`reference.docx` defines `Title`, `Subtitle`, `Abstract`, `Figure`, `Table` and
the heading styles. To change the look of the output, edit those styles in
`reference.docx` (or supply your own template with `-t`) and run the converter
again. The script copies the chosen template into the paper folder as
`reference.docx` before calling pandoc.

## Feedback

Comments, bug reports and suggestions are welcome. Please open an issue on GitHub
or send an email to thaitm@huflit.edu.vn.

## License

Released under the [MIT License](LICENSE). You may use, modify and redistribute
the code freely, including in your own projects, as long as you keep the
copyright and license notice.

