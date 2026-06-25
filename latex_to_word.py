# -*- coding: utf-8 -*-
"""
latex_to_word.py
================
Convert a LaTeX paper to Microsoft Word (.docx) with pandoc while preserving the
text, the mathematics and the bibliography. The script handles the cases pandoc
does not cover on its own:

  - Inline every \\input{...} so the whole paper is processed as one file.
  - Convert the `algorithm`/`algorithmic` blocks into numbered pseudocode while
    keeping inline math as real Word equations.
  - Resolve cross references: \\ref / \\eqref are replaced by the numbers LaTeX
    already computed (read from the .aux file).
  - Number displayed equations on the right margin, the usual Word layout.
  - Replace \\Bigl \\Bigr \\big| ... by \\left \\right so the formulas become
    Word equations (pandoc's math reader does not support the \\big* family).
  - Reference figures by their bare file name (handles \\graphicspath) and embed
    the PNG version of each figure.
  - Build a numbered bibliography:
        * Citations [1], [2], ... are numbered in ORDER OF APPEARANCE of \\cite.
        * The reference text is taken from the .bbl produced by IEEEtran.bst.
        * A DOI is appended to each entry (read from the `doi` field in the .bib).
        * With --year-after-authors, the year is moved into "(year)" right after
          the last author (author-year style, e.g. for HJS).
  - Apply the reference.docx Word template when present and map the auto-generated
    paragraph styles (figure/table captions, author, date) to the template styles.
  - Print a review checklist at the end: what to verify and fix by hand in Word
    (missing DOIs, unresolved references, theorem numbering, layout, ...).

Because the bibliography is embedded directly, pandoc --citeproc is not needed.
This requires the .bbl file to exist (run LaTeX + BibTeX once to generate it).
If no .bbl is found, the script falls back to pandoc --citeproc with the .bib.

Usage
-----
    python latex_to_word.py [MAIN.tex] [-d PAPER_DIR] [-t TEMPLATE.docx] [-o OUT.docx]

      MAIN.tex        main .tex file (default: main.tex)
      -d, --paper-dir folder holding the LaTeX sources (default: current folder)
      -t, --template  Word template; it is copied into the paper folder as
                      reference.docx and used as the pandoc style template
      -o, --output    output .docx (default: <MAIN>.docx inside the paper folder)
      --year-after-authors
                      put the year in "(year)" after the last author (HJS style)

    The settings can also be passed through the L2W_MAIN, L2W_DIR, L2W_TEMPLATE
    and L2W_YEAR_AFTER_AUTHORS environment variables (useful to set them once for
    a whole session instead of repeating the flags on every call).

    Example: pick the LaTeX folder and the template, write next to the sources:
        python latex_to_word.py main_EN.tex -d ./paper -t ./my_style.docx

Running on Google Colab
-----------------------
    # Cell 1: install a recent pandoc (the apt version on Colab is too old and
    #         lacks --citeproc)
    !wget -q https://github.com/jgm/pandoc/releases/download/3.1.13/pandoc-3.1.13-1-amd64.deb -O /tmp/pandoc.deb
    !dpkg -i /tmp/pandoc.deb

    # Cell 2: run the conversion, choosing the folder and the template
    from google.colab import drive; drive.mount('/content/drive')
    base = '/content/drive/MyDrive/.../Latex2Word'              # adjust this path
    !python "{base}/latex_to_word.py" main_EN.tex \
            -d "{base}/paper" -t "{base}/reference.docx"

Running locally (needs pandoc >= 2.11)
    python3 latex_to_word.py main_EN.tex -d paper -t reference.docx

Notes on fidelity
- Mathematics (including displayed equations), tables, figures and the
  bibliography convert well.
- Custom theorem-like environments (Definition/Lemma/Theorem ...) become plain
  paragraphs; their label and number ("Definition 1.") may need a manual touch up
  in Word.

Feedback: comments and suggestions are welcome at thaitm@huflit.edu.vn
"""
import argparse
import glob
import os
import re
import shutil
import subprocess
import sys
import zipfile


def resolve_config(argv=None):
    """Read the configuration from the command line, with environment variables as
    a fallback. Returns (paper_dir, main_tex, out_path, template).

    paper_dir : folder holding the LaTeX sources (the .tex, .bbl, .bib, figures/).
    main_tex  : name of the main .tex file inside paper_dir.
    out_path  : where to write the .docx.
    template  : optional Word template; it is copied into paper_dir as reference.docx.
    year_after_authors : place the year in "(year)" after the last author.
    """
    p = argparse.ArgumentParser(
        prog="latex_to_word.py",
        description="Convert a LaTeX paper to Microsoft Word (.docx) with pandoc.")
    p.add_argument("main", nargs="?", default=None,
                   help="main .tex file (default: $L2W_MAIN, otherwise main.tex)")
    p.add_argument("-d", "--paper-dir", dest="paper_dir", default=None,
                   help="folder holding the LaTeX sources "
                        "(default: $L2W_DIR, otherwise the current folder)")
    p.add_argument("-t", "--template", default=None,
                   help="Word template .docx; it is copied into the paper folder as "
                        "reference.docx and used by pandoc (default: $L2W_TEMPLATE, "
                        "otherwise reference.docx already in the paper folder)")
    p.add_argument("-o", "--output", default=None,
                   help="output .docx path (default: <main>.docx inside the paper folder)")
    p.add_argument("--year-after-authors", dest="year_after_authors", action="store_true",
                   help="put the publication year in parentheses after the last author "
                        "(author-year style, e.g. for HJS) instead of at the end of the entry")
    args = p.parse_args(argv)

    paper_dir = os.path.abspath(os.path.expanduser(
        args.paper_dir or os.environ.get("L2W_DIR", ".")))
    main_tex = args.main or os.environ.get("L2W_MAIN", "main.tex")
    if not main_tex.endswith(".tex"):
        main_tex += ".tex"
    stem = os.path.splitext(os.path.basename(main_tex))[0]
    out_path = (os.path.abspath(os.path.expanduser(args.output))
                if args.output else os.path.join(paper_dir, stem + ".docx"))
    template = args.template or os.environ.get("L2W_TEMPLATE")
    year_after_authors = args.year_after_authors or (
        os.environ.get("L2W_YEAR_AFTER_AUTHORS", "").lower() in ("1", "true", "yes", "on"))
    return paper_dir, main_tex, out_path, template, year_after_authors


def have_pandoc():
    return shutil.which("pandoc") is not None


def pandoc_version():
    """Return pandoc's (major, minor) version, or (0, 0) if it cannot be read."""
    try:
        out = subprocess.run(["pandoc", "--version"], capture_output=True, text=True).stdout
        m = re.search(r"pandoc\s+v?([0-9]+)\.([0-9]+)", out)
        if m:
            return (int(m.group(1)), int(m.group(2)))
    except Exception:
        pass
    return (0, 0)


def cite_args(base):
    """Fallback citation handling, used only when no .bbl file is available.
    pandoc >= 2.11 uses --citeproc; older versions use the pandoc-citeproc filter."""
    bib = _find_bib(base)
    if not bib:
        print("  [!] no .bib file found - skipping the bibliography.")
        return []
    ver = pandoc_version()
    if ver >= (2, 11):
        return ["--citeproc", "--bibliography", bib]
    if shutil.which("pandoc-citeproc"):
        print(f"  (pandoc {ver[0]}.{ver[1]} is old - using the pandoc-citeproc filter)")
        return ["--filter", "pandoc-citeproc", "--bibliography", bib]
    print(f"  [!] pandoc {ver[0]}.{ver[1]} is too old and pandoc-citeproc is missing - "
          "the bibliography will NOT be generated.")
    return ["--bibliography", bib]


def flatten_inputs(main_path, base):
    """Expand every \\input{...} command into its file content so that pandoc
    processes the whole document at once."""
    with open(main_path, encoding="utf-8") as f:
        text = f.read()

    def repl(m):
        name = m.group(1).strip()
        if not name.endswith(".tex"):
            name += ".tex"
        p = os.path.join(base, name)
        if os.path.exists(p):
            with open(p, encoding="utf-8") as g:
                return "\n" + g.read() + "\n"
        return m.group(0)
    # Repeat a few times in case of nested \input commands.
    for _ in range(5):
        new = re.sub(r"\\input\s*\{([^}]+)\}", repl, text)
        if new == text:
            break
        text = new
    return text


# =========================================================================
#  IEEE bibliography: numbered in order of appearance, with a DOI per entry
# =========================================================================
def _find_bib(base):
    cands = glob.glob(os.path.join(base, "*.bib"))
    return cands[0] if cands else None


def _find_bbl(base, main):
    """Prefer the .bbl that matches the main file name, then any .bbl in base."""
    stem = os.path.splitext(main)[0]
    p = os.path.join(base, stem + ".bbl")
    if os.path.exists(p):
        return p
    cands = glob.glob(os.path.join(base, "*.bbl"))
    return cands[0] if cands else None


def parse_bib_meta(bib_path):
    """Return {citekey: {"doi": ..., "year": ...}} from the .bib file.
    A field that is absent is stored as None."""
    meta = {}
    if not bib_path or not os.path.exists(bib_path):
        return meta
    text = open(bib_path, encoding="utf-8").read()
    # Split into individual @type{key, ...} entries.
    for m in re.finditer(r"@\w+\s*\{\s*([^,\s]+)\s*,(.*?)(?=\n@|\Z)", text, re.S):
        key = m.group(1).strip()
        body = m.group(2)
        dm = re.search(r"\bdoi\s*=\s*[{\"]\s*([^}\"]+?)\s*[}\"]", body, re.I)
        ym = re.search(r"\byear\s*=\s*[{\"]?\s*(\d{4})", body, re.I)
        meta[key] = {"doi": dm.group(1).strip() if dm else None,
                     "year": ym.group(1) if ym else None}
    return meta


def _year_after_authors(body, year):
    """Move the publication year into "(year)" right after the last author.

    IEEEtran prints the year near the end of the entry; some journals (e.g. the
    HJS author-year style) want it in parentheses after the author list, just
    before the title. The author list is the text before the opening title quote
    (``), and the year is removed from its original position at the end."""
    if not year:
        return body
    pos = body.find("``")                    # opening quote of the title
    if pos == -1:
        return body                          # no quoted title: leave the entry as is
    authors = body[:pos].rstrip()
    if authors.endswith(","):
        authors = authors[:-1].rstrip()
    rest = body[pos:]
    # Remove the trailing ", <year>" (take the last match: the venue name may also
    # contain the year, e.g. "2016 IEEE International Conference ...").
    found = list(re.finditer(r",\s*\b" + re.escape(year) + r"\b", rest))
    if found:
        m = found[-1]
        rest = rest[:m.start()] + rest[m.end():]
    return f"{authors} ({year}), {rest}"


def parse_bbl_entries(bbl_path):
    """Return {citekey: cleaned_latex} from an IEEEtran .bbl file.

    Commands that pandoc understands are kept (\\emph, \\url, \\texttt, ``...'',
    ~, --). Only the bbl-specific spacing macros are stripped (\\hskip...\\relax,
    \\BIBentry..., \\newblock, \\providecommand, \\BIBdecl)."""
    entries = {}
    if not bbl_path or not os.path.exists(bbl_path):
        return entries
    text = open(bbl_path, encoding="utf-8").read()
    # Keep only the body of the thebibliography environment.
    mb = re.search(r"\\begin\{thebibliography\}.*?\n(.*)\\end\{thebibliography\}", text, re.S)
    body = mb.group(1) if mb else text
    pat = re.compile(r"\\bibitem(?:\[[^\]]*\])?\s*\{([^}]+)\}(.*?)(?=\\bibitem|\Z)", re.S)
    for m in pat.finditer(body):
        key = m.group(1).strip()
        entries[key] = _clean_bbl_entry(m.group(2))
    return entries


def _clean_bbl_entry(s):
    """Clean a single \\bibitem so pandoc can render it with the right formatting."""
    # Drop IEEEtran inter-word glue: \hskip <dimens> \relax
    s = re.sub(r"\\hskip\b.*?\\relax", " ", s, flags=re.S)
    # Drop the bbl-specific spacing / foreign-language macros.
    s = re.sub(r"\\BIBentry[A-Za-z]*", " ", s)
    s = re.sub(r"\\BIBforeignlanguage\b", " ", s)
    s = re.sub(r"\\providecommand\b", " ", s)
    s = s.replace("\\newblock", " ").replace("\\BIBdecl", " ")
    # Drop any remaining \relax tokens.
    s = s.replace("\\relax", " ")
    # Collapse whitespace and line breaks into single spaces.
    s = re.sub(r"\s+", " ", s).strip()
    # Remove stray spaces before punctuation.
    s = re.sub(r"\s+([.,;])", r"\1", s)
    return s


def _fmt_cite_numbers(nums):
    """Format a citation group IEEE style: [1], [3]; a run of >= 3 becomes [1]-[3]."""
    nums = sorted(set(nums))
    out, i = [], 0
    while i < len(nums):
        j = i
        while j + 1 < len(nums) and nums[j + 1] == nums[j] + 1:
            j += 1
        if j - i >= 2:                      # at least three consecutive numbers
            out.append(f"[{nums[i]}]–[{nums[j]}]")
        else:
            out.extend(f"[{n}]" for n in nums[i:j + 1])
        i = j + 1
    return ", ".join(out)


def inject_ieee_references(text, base, main, year_after_authors=False):
    """Replace each \\cite{...} by its number(s) in order of appearance and append
    a numbered References section (with a DOI per entry).
    Returns (new_text, embedded?, missing_doi, missing_entry).

    When year_after_authors is True, the publication year is moved into "(year)"
    right after the last author (author-year style, e.g. for HJS).

    If no .bbl file is found, the text is returned unchanged together with False so
    main() can fall back to pandoc --citeproc."""
    bbl_path = _find_bbl(base, main)
    if not bbl_path:
        print("  [!] no .bbl file found - will try pandoc --citeproc with the .bib instead.")
        return text, False, [], []

    bbl = parse_bbl_entries(bbl_path)
    meta = parse_bib_meta(_find_bib(base))

    # 1) Assign a number to each key in ORDER OF APPEARANCE of \cite.
    order, cite_map = [], {}
    for m in re.finditer(r"\\cite\s*\{([^}]*)\}", text):
        for k in m.group(1).split(","):
            k = k.strip()
            if k and k not in cite_map:
                order.append(k)
                cite_map[k] = len(order)        # 1-based
    if not order:
        print("  [!] no \\cite found - skipping the bibliography.")
        return text, False, [], []

    # 2) Replace every \cite{...} by its bracketed number(s), IEEE style.
    def cite_repl(m):
        nums = [cite_map[k.strip()] for k in m.group(1).split(",") if k.strip() in cite_map]
        return _fmt_cite_numbers(nums) if nums else m.group(0)
    text = re.sub(r"\\cite\s*\{([^}]*)\}", cite_repl, text)

    # 3) Build the References section in assigned-number order. The @@NONUM@@ marker
    #    tells postprocess_docx to suppress numbering on this heading (the template
    #    auto-numbers Heading 1, but "References" should stay unnumbered).
    lines = ["\\section*{@@NONUM@@References}\n"]
    missing_entry, missing_doi = [], []
    for key in order:
        n = cite_map[key]
        body = bbl.get(key)
        if body is None:
            missing_entry.append(key)
            body = f"[missing \\bibitem for key {key} in the .bbl]"
        info = meta.get(key, {})
        if year_after_authors:
            body = _year_after_authors(body, info.get("year"))
        doi = info.get("doi")
        if doi:
            body = body.rstrip()
            if not body.endswith("."):
                body += "."
            body += f" DOI: \\url{{https://doi.org/{doi}}}"
        else:
            missing_doi.append(key)
        # Each reference is its own paragraph (blank line) so pandoc splits them.
        lines.append(f"[{n}] {body}\n")
    refs_block = "\n" + "\n".join(lines) + "\n"

    # 4) Remove the old bibliography commands and insert References before \end{document}.
    text = re.sub(r"\\bibliographystyle\s*\{[^}]*\}", "", text)
    text = re.sub(r"\\bibliography\s*\{[^}]*\}", "", text)
    text = re.sub(r"\\begin\{thebibliography\}.*?\\end\{thebibliography\}", "", text, flags=re.S)
    if "\\end{document}" in text:
        text = text.replace("\\end{document}", refs_block + "\n\\end{document}", 1)
    else:
        text += refs_block

    print(f"  (References: {len(order)} entries inserted, numbered in order of appearance)")
    if missing_entry:
        print(f"  [!] no .bbl entry found for: {', '.join(missing_entry)}")
    if missing_doi:
        print(f"  [!] no DOI for: {', '.join(missing_doi)} (left as is, no DOI invented)")
    return text, True, missing_doi, missing_entry


# =========================================================================
#  Cross references: \ref / \eqref -> numbers (read from the .aux file)
#  Displayed equations get a right-aligned number as well.
# =========================================================================
def parse_aux_labels(base, main):
    """Return {label: printed_number} read from the .aux file, where LaTeX has
    already resolved every reference number."""
    labels = {}
    stem = os.path.splitext(main)[0]
    cands = [os.path.join(base, stem + ".aux")] + sorted(glob.glob(os.path.join(base, "*.aux")))
    for aux in cands:
        if not os.path.exists(aux):
            continue
        txt = open(aux, encoding="utf-8", errors="ignore").read()
        for m in re.finditer(r"\\newlabel\{([^}]+)\}\{\{([^{}]*)\}", txt):
            key = m.group(1)
            val = re.sub(r"\\[a-zA-Z@]+\s*", "", m.group(2)).strip()  # drop \relax etc.
            if val and key not in labels:
                labels[key] = val
        if labels:
            break
    return labels


def resolve_refs(text, labels):
    """Replace \\ref{k} by its number and \\eqref{k} by (number), using the .aux map."""
    missing = set()

    def rref(m):
        k = m.group(1).strip()
        if k in labels:
            return labels[k]
        missing.add(k)
        return "??"

    def reqref(m):
        k = m.group(1).strip()
        if k in labels:
            return "(" + labels[k] + ")"
        missing.add(k)
        return "(??)"

    text = re.sub(r"\\eqref\s*\{([^}]*)\}", reqref, text)
    text = re.sub(r"\\(?:auto|c|C)?ref\s*\{([^}]*)\}", rref, text)
    if missing:
        print(f"  [!] could not resolve these labels (missing from .aux): {', '.join(sorted(missing))}")
    else:
        print("  (replaced every \\ref/\\eqref with its number from the .aux)")
    return text, sorted(missing)


def number_equations(text, labels):
    """Number each displayed equation the way Word does it natively (no table).

    The equation is emitted as inline display math followed by a marker
    (@@EQ:n@@). pandoc keeps it in a single paragraph; postprocess_docx then turns
    the marker into a right-aligned tab + "(n)", which is exactly how Word stores a
    "#"-numbered equation: the formula centered and the number at the right margin.
    (pandoc/texmath ignores \\tag, so it cannot be used here.)"""
    def repl(m):
        body = m.group(1)
        lm = re.search(r"\\label\{([^}]*)\}", body)
        num = labels.get(lm.group(1).strip()) if lm else None
        if lm:
            body = body.replace(lm.group(0), "")
        body = re.sub(r"\s+", " ", body).strip()
        if not num:                                   # unlabeled equation: keep as is
            return "\\begin{equation}" + m.group(1) + "\\end{equation}"
        return "\n\n$\\displaystyle " + body + "$ @@EQ:" + num + "@@\n\n"
    return re.sub(r"\\begin\{equation\}(.*?)\\end\{equation\}", repl, text, flags=re.S)


# =========================================================================
#  Pseudocode (algorithmic): keep inline math as equations, number the lines
#  like algpseudocode, and honor the [noend] option (no "end if/for" lines).
# =========================================================================
def _alg_inline(s):
    """Normalize the content of one pseudocode line into LaTeX for pandoc:
    \\Call{f}{x} becomes f(x) and a mid-line \\Return becomes "return"; $...$,
    \\textbf, \\textsc and \\emph are kept untouched (pandoc renders the math and
    the bold/small-caps text itself)."""
    s = re.sub(r"\\Return\b", "return", s)     # \Return used inside a \State line
    out, i = [], 0
    while i < len(s):
        j = s.find("\\Call", i)
        if j == -1:
            out.append(s[i:])
            break
        out.append(s[i:j])
        k = s.find("{", j)
        if k == -1:
            out.append(s[j:])
            break
        A, a2 = _grab_brace(s, k)
        k2 = s.find("{", a2)
        if k2 != -1 and s[a2:k2].strip() == "":
            B, b2 = _grab_brace(s, k2)
        else:
            B, b2 = "", a2
        out.append(_alg_inline(A) + "(" + _alg_inline(B) + ")")
        i = b2
    return re.sub(r"\s+", " ", "".join(out)).strip()


def _render_rest(rest):
    """Render a statement written after \\If/\\Else on the same line (e.g. then return ...)."""
    rest = rest.strip()
    if rest.startswith("\\Return"):
        return ("return " + _alg_inline(rest[len("\\Return"):])).strip()
    if rest.startswith("\\State"):
        return _alg_inline(rest[len("\\State"):])
    return _alg_inline(rest)


def _grab_brace(s, i):
    """Given s[i] == '{', return (balanced content, index just after the matching '}')."""
    depth = 0
    for j in range(i, len(s)):
        if s[j] == "{":
            depth += 1
        elif s[j] == "}":
            depth -= 1
            if depth == 0:
                return s[i + 1:j], j + 1
    return s[i + 1:], len(s)


def _split_cmd(line, cmd):
    """Split "\\cmd{ARG}REST" with a brace-balanced ARG. Return (arg, rest) or None."""
    if not line.startswith("\\" + cmd):
        return None
    k = line.find("{", len(cmd) + 1)
    if k == -1:
        return ("", line[len(cmd) + 1:].strip())
    arg, after = _grab_brace(line, k)
    return (arg, line[after:].strip())


def pseudocode_to_lines(body):
    """Return a list of (number_label, indent_level, latex_content) for each line.

    Line numbering matches algpseudocode: only State/If/ElsIf/Else/For/ForAll/
    While/Return lines are numbered; Require/Ensure become Input/Output (no number);
    the [noend] option drops the End* lines. Inline math ($...$) is kept so pandoc
    builds real equations."""
    lines, indent, num = [], 0, 0

    def emit(text, level, numbered=True):
        nonlocal num
        label = ""
        if numbered:
            num += 1
            label = f"{num}:"
        lines.append((label, max(level, 0), text))

    for raw in body.split("\n"):
        line = raw.strip()
        if not line or line.startswith("\\begin") or line.startswith("\\end"):
            continue
        # Pull out a trailing \Comment{...} and render it as " > ..." text.
        comment = ""
        ci = line.find("\\Comment{")
        if ci != -1:
            arg, after = _grab_brace(line, line.find("{", ci))
            comment = " ▷ " + _alg_inline(arg)
            line = (line[:ci] + line[after:]).strip()
        if line.startswith("\\Require"):
            emit("\\textbf{Input:} " + _alg_inline(line[len("\\Require"):]), indent, False)
        elif line.startswith("\\Ensure"):
            emit("\\textbf{Output:} " + _alg_inline(line[len("\\Ensure"):]), indent, False)
        elif line.startswith("\\State"):
            emit(_alg_inline(line[len("\\State"):]) + comment, indent)
        elif line.startswith("\\ElsIf"):
            arg, rest = _split_cmd(line, "ElsIf")
            txt = "else if " + _alg_inline(arg) + " then"
            if rest:
                txt += " " + _render_rest(rest)
            emit(txt + comment, indent - 1)
        elif line.startswith("\\Else"):
            arg, rest = _split_cmd(line, "Else")
            txt = "else"
            if rest:
                txt += " " + _render_rest(rest)
            emit(txt + comment, indent - 1)
        elif (line.startswith("\\EndIf") or line.startswith("\\EndFor")
              or line.startswith("\\EndWhile") or line.startswith("\\EndLoop")):
            indent = max(0, indent - 1)         # [noend]: do not print an end line
        elif line.startswith("\\If"):
            arg, rest = _split_cmd(line, "If")
            txt = "if " + _alg_inline(arg) + " then"
            if rest:
                txt += " " + _render_rest(rest)
            emit(txt + comment, indent)
            indent += 1
        elif line.startswith("\\ForAll"):
            arg, rest = _split_cmd(line, "ForAll")
            txt = "for all " + _alg_inline(arg) + " do" + (" " + _alg_inline(rest) if rest else "")
            emit(txt + comment, indent); indent += 1
        elif line.startswith("\\For"):
            arg, rest = _split_cmd(line, "For")
            txt = "for " + _alg_inline(arg) + " do" + (" " + _alg_inline(rest) if rest else "")
            emit(txt + comment, indent); indent += 1
        elif line.startswith("\\While"):
            arg, rest = _split_cmd(line, "While")
            txt = "while " + _alg_inline(arg) + " do" + (" " + _alg_inline(rest) if rest else "")
            emit(txt + comment, indent); indent += 1
        elif line.startswith("\\Return"):
            emit("return " + _alg_inline(line[len("\\Return"):]) + comment, indent)
        else:
            emit(_alg_inline(line) + comment, indent)
    return lines


def _format_algorithm(num, caption, lines):
    """Render one algorithm as a left-aligned header (the "Algorithm N" heading plus
    the Input/Output lines, flush left, with a rule above and below) followed by a
    borderless two-column table of numbered steps (line number | code), one row per
    line.

    A table is used for the steps because it guarantees that every line stays on its
    own line in Word, Google Docs and LibreOffice, while inline math is rendered as a
    real equation in each cell. The @@ALG...@@ markers tell postprocess_docx where to
    draw the borders and which table to keep left-aligned; they are then removed."""
    header = [f"\\textbf{{Algorithm {num}: {caption}}}"]
    header += [text for (label, level, text) in lines if not label]   # Input: / Output:
    steps = [(label, level, text) for (label, level, text) in lines if label]

    paras = []
    for i, h in enumerate(header):
        mark = ("@@ALGTOP@@" if i == 0 else "") + ("@@ALGBOT@@" if i == len(header) - 1 else "")
        # No space between the marker and the text, so no stray space is left after
        # the marker is stripped in post-processing.
        paras.append("\\noindent " + mark + h)
    header_latex = "\n\n".join(paras)

    rows = []
    for j, (label, level, text) in enumerate(steps):
        indent = "~" * (2 * level)                      # nested blocks: non-breaking spaces
        cell = (indent + text).replace("&", "\\&")      # '&' is the column separator
        tag = "@@ALGTBL@@" if j == 0 else ""            # marks the table as an algorithm
        rows.append(f"{tag}{label} & {cell}")
    table = ("\\begin{longtable}[]{@{}p{0.07\\linewidth}p{0.9\\linewidth}@{}}\n"
             + " \\\\\n".join(rows) + " \\\\\n"
             + "\\end{longtable}")
    return "\n\n" + header_latex + "\n\n" + table + "\n\n"


def convert_algorithms(text, labels):
    """Replace each algorithm block by an "Algorithm N: ..." heading followed by the
    pseudocode with equations preserved. N comes from the block \\label via the .aux."""
    def repl(m):
        blk = m.group(0)
        caption = ""
        ci = blk.find("\\caption{")
        if ci != -1:
            caption, _ = _grab_brace(blk, blk.find("{", ci))
        caption = re.sub(r"\s+", " ", caption).strip()
        lm = re.search(r"\\label\{([^}]*)\}", blk)
        num = labels.get(lm.group(1).strip(), "?") if lm else "?"
        body = re.search(r"\\begin\{algorithmic\}(?:\[[^\]]*\])?(.*?)\\end\{algorithmic\}", blk, re.S)
        if not body:
            return blk
        return _format_algorithm(num, caption, pseudocode_to_lines(body.group(1)))
    return re.sub(r"\\begin\{algorithm\}.*?\\end\{algorithm\}", repl, text, flags=re.S)


_RULE = 'w:val="single" w:sz="6" w:space="0" w:color="auto"'


def _insert_in_tblpr(tbl, xml):
    """Insert xml into the table's tblPr, after <w:tblW/> (a safe OOXML position)."""
    new = re.sub(r"(<w:tblW\b[^>]*/>)", r"\1" + xml, tbl, count=1)
    return new if new != tbl else tbl.replace("</w:tblPr>", xml + "</w:tblPr>", 1)


def _header_row_rule(tbl):
    """Add a bottom rule to every cell of the table's first row (the header rule)."""
    fr = re.search(r"<w:tr\b.*?</w:tr>", tbl, re.S)
    if not fr:
        return tbl
    row = fr.group(0)
    tcb = f'<w:tcBorders><w:bottom {_RULE}/></w:tcBorders>'

    def add(cell):
        c = cell.group(0)
        if "<w:tcBorders>" in c:
            return c
        if "<w:tcPr>" in c:
            return c.replace("</w:tcPr>", tcb + "</w:tcPr>", 1)
        return re.sub(r"(<w:tc>)", r"\1<w:tcPr>" + tcb + "</w:tcPr>", c, count=1)
    new_row = re.sub(r"<w:tc>.*?</w:tc>", add, row, flags=re.S)
    return tbl.replace(row, new_row, 1)


def _center_table(m):
    """Style one table. Algorithm tables (marked @@ALGTBL@@) stay left-aligned with a
    single closing rule. Content tables are centered and get booktabs-style rules:
    a top rule, a rule under the header row, and a bottom rule. pandoc left-aligns
    tables, ignores \\centering and draws no rules, so we set all of this here."""
    tbl = m.group(0)
    if "@@ALGTBL@@" in tbl:                      # algorithm body
        tbl = tbl.replace("@@ALGTBL@@", "")
        # A cell that contains ONLY a formula can be wrapped by pandoc in a centered
        # display equation (m:oMathPara) - this happens on some pandoc builds (e.g.
        # amd64 on Colab) and makes the first, formula-only line look centered.
        # Unwrap it to inline math so it follows the cell's left alignment.
        tbl = re.sub(r"<m:oMathPara>(?:<m:oMathParaPr>.*?</m:oMathParaPr>)?"
                     r"(<m:oMath>.*?</m:oMath>)</m:oMathPara>", r"\1", tbl, flags=re.S)
        # Force LEFT (the template "Table" style is centered) and close the box.
        extra = '<w:jc w:val="left"/>'
        if "<w:tblBorders>" not in tbl:
            extra += f'<w:tblBorders><w:bottom {_RULE}/></w:tblBorders>'
        return _insert_in_tblpr(tbl, extra)

    # Content table: center it, add top/bottom rules and a header rule.
    extra = ""
    if "<w:jc " not in re.search(r"<w:tblPr>.*?</w:tblPr>", tbl, re.S).group(0):
        extra += '<w:jc w:val="center"/>'
    if "<w:tblBorders>" not in tbl:
        extra += f'<w:tblBorders><w:top {_RULE}/><w:bottom {_RULE}/></w:tblBorders>'
    if extra:
        tbl = _insert_in_tblpr(tbl, extra)
    return _header_row_rule(tbl)


def _suppress_heading_numbering(s):
    """Suppress automatic numbering on any heading marked @@NONUM@@ (the generated
    References heading): the template numbers Heading 1, but References should not be
    numbered. numId=0 removes the paragraph from the heading's number list."""
    nonum = '<w:numPr><w:ilvl w:val="0"/><w:numId w:val="0"/></w:numPr>'

    def fix(m):
        p = m.group(0)
        if "@@NONUM@@" not in p:
            return p
        if "<w:pStyle " in p:
            p = re.sub(r"(<w:pStyle\b[^>]*/>)", r"\1" + nonum, p, count=1)
        elif "<w:pPr>" in p:
            p = p.replace("<w:pPr>", "<w:pPr>" + nonum, 1)
        else:
            p = re.sub(r"(<w:p\b[^>]*>)", r"\1<w:pPr>" + nonum + "</w:pPr>", p, count=1)
        return p.replace("@@NONUM@@", "")

    return re.sub(r"<w:p\b.*?</w:p>", fix, s, flags=re.S)


def _style_algorithm_borders(s):
    """Add a top rule above the algorithm heading (@@ALGTOP@@) and a bottom rule
    below the last header line (@@ALGBOT@@), then remove the markers. Returns
    (new_xml, count_of_algorithms)."""
    top = '<w:top w:val="single" w:sz="6" w:space="2" w:color="auto"/>'
    bottom = '<w:bottom w:val="single" w:sz="6" w:space="2" w:color="auto"/>'
    count = [0]

    def fix(m):
        p = m.group(0)
        has_top, has_bot = "@@ALGTOP@@" in p, "@@ALGBOT@@" in p
        if not (has_top or has_bot):
            return p
        if has_top:
            count[0] += 1
        pbdr = "<w:pBdr>" + (top if has_top else "") + (bottom if has_bot else "") + "</w:pBdr>"
        if "<w:pStyle " in p:                    # OOXML order: pBdr right after pStyle is valid
            p = re.sub(r"(<w:pStyle\b[^>]*/>)", r"\1" + pbdr, p, count=1)
        elif "<w:pPr>" in p:
            p = p.replace("<w:pPr>", "<w:pPr>" + pbdr, 1)
        else:
            p = re.sub(r"(<w:p\b[^>]*>)", r"\1<w:pPr>" + pbdr + "</w:pPr>", p, count=1)
        return p.replace("@@ALGTOP@@", "").replace("@@ALGBOT@@", "")

    return re.sub(r"<w:p\b.*?</w:p>", fix, s, flags=re.S), count[0]


def _text_width_twips(s):
    """Usable text width (page width minus left/right margins) in twips, from the
    section properties. Falls back to a sensible A4 default."""
    sect = re.search(r"<w:sectPr.*?</w:sectPr>", s, re.S)
    if sect:
        pg = re.search(r'<w:pgSz\b[^>]*\bw:w="(\d+)"', sect.group(0))
        mar = re.search(r"<w:pgMar\b[^>]*/>", sect.group(0))
        if pg and mar:
            left = re.search(r'\bw:left="(\d+)"', mar.group(0))
            right = re.search(r'\bw:right="(\d+)"', mar.group(0))
            if left and right:
                return int(pg.group(1)) - int(left.group(1)) - int(right.group(1))
    return 9026     # A4 with ~1 inch margins


def _apply_equation_numbers(s):
    """Turn each @@EQ:n@@ marker into a Word-native numbered equation: the formula
    centered and "(n)" right-aligned, using tab stops (no table). Returns
    (new_xml, count)."""
    width = _text_width_twips(s)
    tabs = (f'<w:tabs><w:tab w:val="center" w:pos="{width // 2}"/>'
            f'<w:tab w:val="right" w:pos="{width}"/></w:tabs>')
    count = [0]

    def fix(p):
        block = p.group(0)
        if "@@EQ:" not in block:
            return block
        count[0] += 1
        # 1) add the tab stops to the paragraph properties (create pPr if missing)
        if "<w:pPr>" in block:
            block = block.replace("</w:pPr>", tabs + "</w:pPr>", 1)
        else:
            block = re.sub(r"(<w:p\b[^>]*>)", r"\1<w:pPr>" + tabs + "</w:pPr>", block, count=1)
        # 2) a leading center tab so the formula is centered on the page
        block = re.sub(r"(</w:pPr>)\s*(<m:oMath>)", r"\1<w:r><w:tab/></w:r>\2", block, count=1)
        # 3) the marker run (and the lone space run before it) -> right tab + "(n)"
        block = re.sub(
            r'(?:<w:r>(?:<w:rPr>.*?</w:rPr>)?<w:t[^>]*>\s*</w:t></w:r>)?'
            r'<w:r>(?:<w:rPr>.*?</w:rPr>)?<w:t[^>]*>\s*@@EQ:(\d+)@@\s*</w:t></w:r>',
            r'<w:r><w:tab/></w:r><w:r><w:t>(\1)</w:t></w:r>', block, count=1, flags=re.S)
        return block

    s = re.sub(r"<w:p\b.*?</w:p>", fix, s, flags=re.S)
    return s, count[0]


def postprocess_docx(path):
    """Adjust the .docx pandoc produced (plain string edits on document.xml):
      - Map auto-generated paragraph styles onto the template styles (these styles
        already exist in reference.docx):
          Figure caption : ImageCaption -> Figure
          Table caption  : TableCaption -> Table
          Author / Date  : Author, Date -> Subtitle (template has no such styles)
      - Add a rule above each algorithm heading and below its Input/Output header.
      - Center content tables, but keep algorithm tables left-aligned (pandoc
        left-aligns tables and ignores \\centering).
      - Turn the @@EQ:n@@ markers into right-aligned equation numbers (no table).
    """
    remap = {"ImageCaption": "Figure", "TableCaption": "Table",
             "Author": "Subtitle", "Date": "Subtitle"}
    tmp = path + ".tmp"
    changed = 0
    centered = 0
    numbered = 0
    algos = 0
    try:
        with zipfile.ZipFile(path) as zin, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                data = zin.read(item.filename)
                if item.filename == "word/document.xml":
                    s = data.decode("utf-8")
                    for a, b in remap.items():
                        c = s.count(f'w:val="{a}"')
                        if c:
                            s = s.replace(f'w:val="{a}"', f'w:val="{b}"'); changed += c
                    s = _suppress_heading_numbering(s)
                    s, algos = _style_algorithm_borders(s)
                    before = s.count('<w:jc w:val="center"')
                    s = re.sub(r"<w:tbl>.*?</w:tbl>", _center_table, s, flags=re.S)
                    centered = s.count('<w:jc w:val="center"') - before
                    s, numbered = _apply_equation_numbers(s)
                    data = s.encode("utf-8")
                zout.writestr(item, data)
        shutil.move(tmp, path)
        if changed:
            print(f"  (post-processing: remapped {changed} paragraphs to Figure/Table/Subtitle)")
        if algos:
            print(f"  (post-processing: framed {algos} algorithm header(s) with top/bottom rules)")
        if centered:
            print(f"  (post-processing: centered {centered} content table(s))")
        if numbered:
            print(f"  (post-processing: numbered {numbered} equation(s) with a right-aligned tab)")
    except Exception as e:
        if os.path.exists(tmp):
            os.remove(tmp)
        print(f"  [!] skipping style post-processing (optional): {e}")


def print_review_checklist(source_text, missing_doi, missing_entry, unresolved_refs,
                           failed_figures, year_after_authors):
    """Print a short list of things the user should check in Word and fix by hand.

    Combines findings from this run (unresolved references, missing DOIs, figures
    that did not embed, ...) with standing reminders about the parts pandoc cannot
    reproduce exactly."""
    found = []
    if unresolved_refs:
        found.append(
            f"{len(unresolved_refs)} cross-reference(s) show as '??' (label not in the .aux): "
            f"{', '.join(unresolved_refs)}. Run LaTeX once to refresh the .aux, then convert again.")
    if missing_entry:
        found.append(
            f"{len(missing_entry)} citation(s) have no .bbl entry and appear as a placeholder: "
            f"{', '.join(missing_entry)}. Re-run BibTeX so the .bbl includes them.")
    if missing_doi:
        found.append(
            f"{len(missing_doi)} reference(s) have no DOI: {', '.join(missing_doi)}. "
            f"Add a 'doi' field in the .bib if one exists, then convert again.")
    if failed_figures:
        found.append(
            f"{len(failed_figures)} figure(s) did not embed: {', '.join(sorted(failed_figures))}. "
            f"Provide a PNG with that name next to the source, or check the path.")

    reminders = []
    if re.search(r"\\begin\{(theorem|lemma|corollary|proposition|definition|remark|proof)\}",
                 source_text):
        reminders.append(
            "Theorem/Lemma/Definition blocks become plain paragraphs. Their label and number "
            "(e.g. \"Theorem 1.\") are not added automatically; add them by hand if your style needs them.")
    reminders.append(
        "Displayed equation numbers use a right-aligned tab (Word's native style): the formula is "
        "centered and the number sits at the right margin. Check the alignment, especially for wide formulas.")
    reminders.append(
        "Pseudocode has a left-aligned header (heading, Input, Output) framed by a top and bottom "
        "rule, with the steps in a two-column table (line number | code). Check the indentation and "
        "that long formulas do not wrap awkwardly.")
    reminders.append(
        "Check figure size and placement, and table column widths; pandoc uses default sizing.")
    reminders.append(
        "Confirm the Title, Subtitle and heading styles look right for your template.")
    if year_after_authors:
        reminders.append(
            "Author-year style is on: confirm each reference shows the year as \"(year)\" right after "
            "the last author.")
    reminders.append(
        "If you change any \\cite or \\ref/\\eqref later, re-run LaTeX + BibTeX so the .aux and .bbl "
        "refresh before converting again.")

    print("\n" + "-" * 72)
    print("REVIEW CHECKLIST - open the .docx in Word and check the following:")
    n = 1
    if found:
        print("\n  From this run:")
        for item in found:
            print(f"   {n}. {item}")
            n += 1
    print("\n  General reminders:")
    for item in reminders:
        print(f"   {n}. {item}")
        n += 1
    print("-" * 72)


def main():
    if not have_pandoc():
        sys.exit("ERROR: pandoc not found. On Colab, install a recent .deb (see the header).\n"
                 "Locally, install the .pkg from https://pandoc.org (no Homebrew needed).")

    base, main_tex, out_path, template, year_after_authors = resolve_config()
    main_path = os.path.join(base, main_tex)
    if not os.path.exists(main_path):
        sys.exit(f"ERROR: cannot find {main_path}")

    # 0) Copy the chosen Word template into the paper folder as reference.docx so
    #    pandoc can use it. This lets the user keep the template anywhere.
    ref_target = os.path.join(base, "reference.docx")
    if template:
        template = os.path.abspath(os.path.expanduser(template))
        if not os.path.exists(template):
            sys.exit(f"ERROR: template not found: {template}")
        if os.path.abspath(template) != os.path.abspath(ref_target):
            shutil.copy(template, ref_target)
            print(f"  (copied template '{os.path.basename(template)}' -> reference.docx)")

    # 1) Inline every \input into a single document.
    text = flatten_inputs(main_path, base)
    source_text = text                      # kept for the review checklist (theorem scan)

    # 2) Bibliography: numbers in order of appearance + DOI per entry.
    text, refs_embedded, missing_doi, missing_entry = inject_ieee_references(
        text, base, main_tex, year_after_authors)

    # 3) Cross references: \ref/\eqref -> numbers, and number displayed equations.
    labels = parse_aux_labels(base, main_tex)
    if not labels:
        print("  [!] no reference numbers found in the .aux - \\ref/\\eqref may stay unresolved.\n"
              "      (run LaTeX once to generate the .aux file)")
    text, unresolved_refs = resolve_refs(text, labels)
    text = number_equations(text, labels)

    # 4) Algorithms: keep formulas as equations, number lines, "Algorithm N: ...".
    text = convert_algorithms(text, labels)

    # 5) pandoc's math reader (texmath) does not support \Bigl \Bigr \big| ...
    #    Map them to \left \right (or drop them) so the formulas become equations.
    for a, b in [("\\Biggl", "\\left"), ("\\Biggr", "\\right"), ("\\biggl", "\\left"), ("\\biggr", "\\right"),
                 ("\\Bigl", "\\left"), ("\\Bigr", "\\right"), ("\\bigl", "\\left"), ("\\bigr", "\\right"),
                 ("\\Bigg", ""), ("\\bigg", ""), ("\\Big", ""), ("\\big", "")]:
        text = text.replace(a, b)
    text = text.replace(".pdf}", ".png}")   # Word embeds PNG more reliably than PDF
    # Drop a couple of PDF-only commands that pandoc does not need.
    text = text.replace("\\resizebox{\\textwidth}{!}{%", "").replace("\\resizebox{\\textwidth}{!}{", "")
    text = text.replace("\\end{tabular}}", "\\end{tabular}")

    # 6) Write a temporary flattened file inside the paper folder so the figures/
    #    paths stay valid.
    flat = os.path.join(base, "_" + os.path.splitext(main_tex)[0] + "_flat.tex")
    with open(flat, "w", encoding="utf-8") as f:
        f.write(text)

    # 7) Warn about any figure that has no PNG version yet.
    fig_dir = os.path.join(base, "figures")
    if os.path.isdir(fig_dir):
        for fn in os.listdir(fig_dir):
            if fn.endswith(".pdf") and not os.path.exists(os.path.join(fig_dir, fn[:-4] + ".png")):
                print(f"  [!] no PNG for {fn} - this figure may not show up in Word.")

    # 8) Run pandoc (out_path comes from resolve_config).
    cmd = [
        "pandoc", flat,
        "--from", "latex",
        "--to", "docx",
    ]
    if not refs_embedded:                       # fallback: let pandoc handle citations
        cmd += cite_args(base)
    # resource-path: the paper folder plus every \graphicspath folder (pandoc does
    # not read \graphicspath), so figures referenced by bare name (e.g. fig1.png)
    # are found.
    res_dirs = [base]
    for g in re.findall(r"\\graphicspath\{((?:\{[^}]*\})+)\}", text):
        for d in re.findall(r"\{([^}]*)\}", g):
            res_dirs.append(os.path.join(base, d))
    if os.path.isdir(os.path.join(base, "figures")):
        res_dirs.append(os.path.join(base, "figures"))
    seen = []
    for d in res_dirs:
        if d not in seen:
            seen.append(d)
    cmd += [
        "--resource-path", os.pathsep.join(seen),
        # No --number-sections: the Word template's heading styles number the
        # headings automatically, so letting pandoc number them too would double up.
        "-o", out_path,
    ]
    if os.path.exists(ref_target):                      # optional Word template
        cmd += ["--reference-doc", ref_target]
        print("  (using the reference.docx Word template)")

    print(">>> running pandoc ...")
    try:
        r = subprocess.run(cmd, capture_output=True, text=True)
    finally:
        if os.path.exists(flat):
            os.remove(flat)
    if r.returncode != 0:
        print(r.stdout); print(r.stderr)
        sys.exit("pandoc failed - see the messages above.")
    postprocess_docx(out_path)
    print(f"DONE -> {out_path}")
    if r.stderr.strip():
        print("pandoc warnings (usually harmless):")
        print(r.stderr.strip()[:1500])

    # Figures pandoc could not embed (reported in its stderr).
    failed_figures = set(re.findall(r"Could not fetch resource ([^:]+):", r.stderr))

    # Final reminder of what the user should verify by hand.
    print_review_checklist(source_text, missing_doi, missing_entry, unresolved_refs,
                           failed_figures, year_after_authors)


if __name__ == "__main__":
    main()
