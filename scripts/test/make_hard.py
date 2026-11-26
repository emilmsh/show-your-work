"""Builds the files the look-up is tested on, each with what real files have
and the demo set has not: text in grouped shapes and tables, soft line breaks,
formatted numbers, merged cells, sheet names with spaces and apostrophes,
text files in three encodings, Markdown marks, HTML entities and inline tags.

    python scripts/test/make_hard.py <out folder>

Writes the files and `cases.json` beside them: every case is a quote, where
it is, and whether the look-up must find it. `check-cases.ps1` runs them.
"""

import json
import sys
from datetime import date
from pathlib import Path

from openpyxl import Workbook
from pptx import Presentation
from pptx.util import Inches

OUT = Path(sys.argv[1])
OUT.mkdir(parents=True, exist_ok=True)
cases: list[dict] = []


def case(kind: str, file: str, quote: str, found: bool = True, page: int = 0, cell: str = "", why: str = "") -> None:
    cases.append({"kind": kind, "file": file, "quote": quote, "cell": cell, "found": found, "page": page, "why": why})


def slides() -> None:
    deck = Presentation()
    body_layout, blank = deck.slide_layouts[1], deck.slide_layouts[6]
    f = "hard.pptx"

    s = deck.slides.add_slide(body_layout)
    s.shapes.title.text = "Feilutbetalinger i fire land"
    s.placeholders[1].text_frame.text = "Sverige anslår feilutbetalingene til 4,1 prosent av utbetalt beløp."
    case("powerpoint", f, "Sverige anslår feilutbetalingene til 4,1 prosent av utbetalt beløp.", page=1, why="plain body, æøå")

    s = deck.slides.add_slide(blank)
    group = s.shapes.add_group_shape()
    box = group.shapes.add_textbox(Inches(1), Inches(1), Inches(4), Inches(1))
    box.text_frame.text = "Grouped label: 37 municipalities took part."
    inner = group.shapes.add_group_shape()
    deep = inner.shapes.add_textbox(Inches(1), Inches(3), Inches(4), Inches(1))
    deep.text_frame.text = "A group inside a group holds 12 of them."
    case("powerpoint", f, "Grouped label: 37 municipalities took part.", page=2, why="text in a grouped shape")
    case("powerpoint", f, "A group inside a group holds 12 of them.", page=2, why="text in a group inside a group")

    s = deck.slides.add_slide(blank)
    table = s.shapes.add_table(3, 2, Inches(1), Inches(1), Inches(6), Inches(2)).table
    table.cell(0, 0).text, table.cell(0, 1).text = "Land", "Andel"
    table.cell(1, 0).text, table.cell(1, 1).text = "Danmark", "2,7 prosent av utbetalingene"
    table.cell(2, 0).text, table.cell(2, 1).text = "Finland", "1,9 prosent"
    case("powerpoint", f, "2,7 prosent av utbetalingene", page=3, why="a table cell")
    case("powerpoint", f, "Danmark 2,7 prosent av utbetalingene", page=3, why="two cells of one row, quoted as a line")

    s = deck.slides.add_slide(body_layout)
    s.shapes.title.text = "Soft breaks"
    frame = s.placeholders[1].text_frame
    frame.text = "The first line ends here"
    frame.paragraphs[0].add_line_break()
    frame.paragraphs[0].add_run().text = "and the second goes on."
    frame.add_paragraph().text = "A new paragraph says the rest."
    frame.add_paragraph().text = "The agency’s own figure is “about a third”."
    case("powerpoint", f, "The first line ends here and the second goes on.", page=4, why="a soft line break in the quote")
    case("powerpoint", f, "and the second goes on. A new paragraph says the rest.", page=4, why="a quote across two paragraphs")
    case("powerpoint", f, "The agency's own figure is \"about a third\".", page=4, why="plain quotation marks for typeset ones")

    s = deck.slides.add_slide(body_layout)
    s.shapes.title.text = "Notes"
    s.placeholders[1].text_frame.text = "See the notes."
    s.notes_slide.notes_text_frame.text = "Say that the Dutch estimate is from 2022."
    case("powerpoint", f, "Say that the Dutch estimate is from 2022.", page=5, why="speaker notes")

    s = deck.slides.add_slide(body_layout)
    s.shapes.title.text = "Twice"
    s.placeholders[1].text_frame.text = "Sverige anslår feilutbetalingene til 4,1 prosent av utbetalt beløp. Again."
    case("powerpoint", f, "Feilutbetalinger i fire land", page=1, why="a title")
    case("powerpoint", f, "Sverige anslår feilutbetalingene til 5,1 prosent", found=False, why="a figure that is not there")
    case("powerpoint", f, "Grouped label: 38 municipalities", found=False, why="a figure that is not there, in a group")
    deck.save(OUT / f)


def workbook() -> None:
    book = Workbook()
    f = "hard.xlsx"
    ws = book.active
    ws.title = "Tall 2024"
    ws["A1"], ws["B1"] = "Størrelse", "Verdi"
    ws["A2"], ws["B2"] = "Andel feil", 0.042
    ws["B2"].number_format = "0.0%"
    ws["A3"], ws["B3"] = "Utbetalt", 1234567
    ws["B3"].number_format = "#,##0"
    ws["A4"], ws["B4"] = "Mottakere", 1987
    ws["A5"], ws["B5"] = "Dato", date(2024, 3, 15)
    ws["B5"].number_format = "yyyy-mm-dd"
    ws["A6"], ws["B6"] = "Anslag", "ca. 40 %"
    ws["A7"], ws["B7"] = "Endring", -12.5
    ws["A8"], ws["B8"] = "Avrundet", 0.04178
    ws["B8"].number_format = "0.0%"
    ws["A9"], ws["B9"] = "Millioner", 9500000
    ws["B9"].number_format = '0.0,," mill."'
    ws["A10"] = "En overskrift over tre kolonner"
    ws.merge_cells("A10:C10")
    ws["A11"], ws["B11"] = "År", 2024
    ws["A12"], ws["B12"] = "Kroner", 4800.5
    ws["B12"].number_format = '#,##0.00 "kr"'
    ws.column_dimensions["A"].width = 30
    ws.column_dimensions["B"].width = 16
    t = "Tall 2024"
    case("excel", f, "4,2 %", cell=f"{t}!B2", why="a percentage, Norwegian form")
    case("excel", f, "4.2%", cell=f"{t}!B2", why="a percentage, English form")
    case("excel", f, "0,042", cell=f"{t}!B2", why="a percentage cited as its fraction")
    case("excel", f, "4,3 %", cell=f"{t}!B2", found=False, why="the wrong percentage")
    case("excel", f, "1 234 567", cell=f"{t}!B3", why="thousands by spaces")
    case("excel", f, "1,234,567", cell=f"{t}!B3", why="thousands by commas")
    case("excel", f, "1.234.567", cell=f"{t}!B3", why="thousands by full stops")
    case("excel", f, "1,987", cell=f"{t}!B4", why="1987 written the English way")
    case("excel", f, "1.987", cell=f"{t}!B4", why="1987 written the continental way")
    case("excel", f, "1,988", cell=f"{t}!B4", found=False, why="the wrong whole number")
    case("excel", f, "2024-03-15", cell=f"{t}!B5", why="a date as the cell shows it")
    case("excel", f, "ca. 40 %", cell=f"{t}!B6", why="a text cell")
    case("excel", f, "40 %", cell=f"{t}!B6", found=False, why="part of a text cell is not the cell")
    case("excel", f, "-12,5", cell=f"{t}!B7", why="a negative number")
    case("excel", f, "−12.5", cell=f"{t}!B7", why="a typeset minus sign")
    case("excel", f, "12,5", cell=f"{t}!B7", found=False, why="the sign left out")
    case("excel", f, "4,2 %", cell=f"{t}!B8", why="cited as the cell shows it, rounded")
    case("excel", f, "4,18 %", cell=f"{t}!B8", why="cited more precisely than the cell shows")
    case("excel", f, "4,3 %", cell=f"{t}!B8", found=False, why="rounded wrongly")
    case("excel", f, "9,5 mill.", cell=f"{t}!B9", why="a format that shows millions")
    case("excel", f, "9 500 000", cell=f"{t}!B9", why="the value under a format that shows millions")
    case("excel", f, "En overskrift over tre kolonner", cell=f"{t}!A10", why="a merged cell by its first cell")
    case("excel", f, "En overskrift over tre kolonner", cell=f"{t}!B10", why="a merged cell by another of its cells")
    case("excel", f, "2024", cell=f"{t}!B11", why="a year")
    case("excel", f, "4 800,50 kr", cell=f"{t}!B12", why="a number with its unit")
    case("excel", f, "4800.5", cell=f"{t}!B12", why="the bare number under a unit")
    case("excel", f, "1987", cell=f"'{t}'!B4", why="a sheet name in quotation marks")
    case("excel", f, "1987", cell=f"{t}!$B$4", why="an absolute address")
    case("excel", f, "1987", cell=f"{t}!B:B", why="found in a column")
    case("excel", f, "1,987", cell=f"{t}!B:B", why="found in a column, English form")
    case("excel", f, "1987", cell=f"{t}!4:4", why="found in a row")
    case("excel", f, "1987", cell=f"{t}!A1:B12", why="found in a block")
    case("excel", f, "1986", cell=f"{t}!A1:B12", found=False, why="not in the block")
    case("excel", f, t, cell=f"{t}!", why="a sheet alone")
    case("excel", f, "1987", cell="B4", why="no sheet named: the first sheet")

    other = book.create_sheet("O'Brien & Co")
    other["A1"], other["B1"] = "Fee", 250
    case("excel", f, "250", cell="'O''Brien & Co'!B1", why="an apostrophe in the sheet name, as Excel writes it")
    case("excel", f, "250", cell="O'Brien & Co!B1", why="an apostrophe in the sheet name, written plainly")

    hidden = book.create_sheet("Skjult")
    hidden["A1"], hidden["B1"] = "Hidden figure", 77
    hidden.sheet_state = "hidden"
    case("excel", f, "77", cell="Skjult!B1", why="a hidden sheet")
    case("excel", f, "1987", cell="Finnes ikke!B4", found=False, why="no such sheet")
    book.save(OUT / f)


def texts() -> None:
    md = "\r\n".join(
        [
            "# Notat om feilutbetalinger",
            "",
            "Skrevet etter møtet 12. november.",
            "",
            "## Funn",
            "",
            "- **Feilutbetalingene** utgjorde 1,2 mrd. kroner i 2023.",
            "- Se [rapporten fra Riksrevisjonen](https://example.org/rapport) for metoden.",
            "- Anslaget er _usikkert_, og `n = 412` er et lite utvalg.",
            "",
            "Sverige bruker innloggingsdata til å anslå uoppgitt samboerskap,",
            "og Danmark bruker registerkobling til det samme.",
            "",
            "| Land | Andel |",
            "| --- | --- |",
            "| Sverige | 4,1 % |",
            "",
            "Siste linje, uten linjeskift etter seg.",
        ]
    )
    (OUT / "hard.md").write_bytes(md.encode("utf-8"))
    case("text", "hard.md", "Skrevet etter møtet 12. november.", page=3, why="æøå in UTF-8 without a mark, CRLF lines")
    case("text", "hard.md", "Feilutbetalingene utgjorde 1,2 mrd. kroner i 2023.", page=7, why="bold marks left out of the quote")
    case("text", "hard.md", "**Feilutbetalingene** utgjorde 1,2 mrd. kroner i 2023.", page=7, why="bold marks kept in the quote")
    case("text", "hard.md", "Se rapporten fra Riksrevisjonen for metoden.", page=8, why="a link quoted by its words")
    case("text", "hard.md", "Anslaget er usikkert, og n = 412 er et lite utvalg.", page=9, why="italics and code marks left out")
    case("text", "hard.md", "anslå uoppgitt samboerskap, og Danmark bruker registerkobling", page=11, why="a quote across a line break")
    case("text", "hard.md", "Sverige | 4,1 %", page=16, why="a table row as written")
    case("text", "hard.md", "Siste linje, uten linjeskift etter seg.", page=18, why="the last line")
    case("text", "hard.md", "utgjorde 1,3 mrd. kroner", found=False, why="a figure that is not there")

    ansi = "Første linje.\r\nBeløpet på 4 800 kr gjelder året 2024.\r\nTredje linje: blåbærsyltetøy.\r\n"
    (OUT / "hard-ansi.txt").write_bytes(ansi.encode("cp1252"))
    case("text", "hard-ansi.txt", "Beløpet på 4 800 kr gjelder året 2024.", page=2, why="æøå in the Windows code page")
    case("text", "hard-ansi.txt", "Tredje linje: blåbærsyltetøy.", page=3, why="æøå in the Windows code page")

    (OUT / "hard-utf16.txt").write_bytes(ansi.encode("utf-16"))
    case("text", "hard-utf16.txt", "Beløpet på 4 800 kr gjelder året 2024.", page=2, why="UTF-16 with its mark")

    (OUT / "hard-bom.txt").write_bytes(b"\xef\xbb\xbf" + ansi.replace("\r\n", "\n").encode("utf-8"))
    case("text", "hard-bom.txt", "Første linje.", page=1, why="UTF-8 with its mark, LF lines")
    case("text", "hard-bom.txt", "Tredje linje: blåbærsyltetøy.", page=3, why="UTF-8 with its mark, LF lines")

    code = "\r\n".join(
        [
            "# Modell for feilutbetalinger",
            "library(dplyr)",
            "",
            "utvalg <- data |>",
            "\tfilter(år >= 2019, !is.na(beløp)) |>",
            "\tmutate(andel = feil / utbetalt)",
            "",
            'm <- feols(andel ~ tiltak | kommune + år, data = utvalg, cluster = ~kommune)',
            "",
        ]
    )
    (OUT / "hard.R").write_bytes(code.encode("utf-8"))
    case("text", "hard.R", "filter(år >= 2019, !is.na(beløp)) |>", page=5, why="code with a tab in front and æøå")
    case("text", "hard.R", "feols(andel ~ tiltak | kommune + år, data = utvalg, cluster = ~kommune)", page=8, why="regex characters in code")
    case("text", "hard.R", "filter(år >= 2019, !is.na(beløp)) |> mutate(andel = feil / utbetalt)", page=5, why="two lines of code as one quote")
    case("text", "hard.R", "filter(år >= 2018", found=False, why="a number that is not there")


def html() -> None:
    page = """<!doctype html>
<html lang="nb"><head><meta charset="utf-8"><title>Tittelen st&aring;r i hodet</title>
<style>p { color: black; } /* skjult stiltekst 111 */</style>
<script>var hemmelig = "skripttekst 222";</script></head>
<body>
<!-- en kommentar med > i seg og ordene kommentartekst 333 -->
<h1>Feilutbetalinger i Norden</h1>
<p>Sverige ansl&aring;r feilutbetalingene til <b>4,1</b>&nbsp;prosent av utbetalt bel&oslash;p.</p>
<p>Danmark bruker register&shy;kobling &mdash; og f&#229;r lavere tall &#x2013; rundt 2,7&#160;prosent.</p>
<p>Et ord med ut<i>hev</i>ing midt i seg, og en lenke til <a href="https://example.org" title="titteltekst 444">rapporten</a>.</p>
<p>F&oslash;rste linje<br>andre linje<br/>tredje linje.</p>
<p>Etatens &laquo;eget&raquo; anslag er &ldquo;en tredel&rdquo;, og 3&nbsp;&lt;&nbsp;5 &amp; 5&nbsp;&gt;&nbsp;3.</p>
<ul><li>F&oslash;rste punkt</li><li>Andre punkt</li></ul>
<table><tr><td>Finland</td><td>1,9 prosent</td></tr></table>
</body></html>
"""
    f = "hard.html"
    (OUT / f).write_bytes(page.encode("utf-8"))
    case("html", f, "Sverige anslår feilutbetalingene til 4,1 prosent av utbetalt beløp.", why="named entities, a tag in the sentence")
    case("html", f, "Danmark bruker registerkobling — og får lavere tall – rundt 2,7 prosent.", why="numbered entities, a soft hyphen, dashes")
    case("html", f, "Et ord med utheving midt i seg", why="a tag inside a word")
    case("html", f, "en lenke til rapporten.", why="a link's words")
    case("html", f, "Første linje andre linje tredje linje.", why="line breaks")
    case("html", f, "Etatens «eget» anslag er \"en tredel\", og 3 < 5 & 5 > 3.", why="quotation marks and escaped signs")
    case("html", f, "Første punkt Andre punkt", why="two list items")
    case("html", f, "Finland 1,9 prosent", why="two table cells")
    case("html", f, "Feilutbetalinger i Norden", why="a heading")
    case("html", f, "skripttekst 222", found=False, why="text of a script is not the page's")
    case("html", f, "skjult stiltekst 111", found=False, why="text of a style sheet is not the page's")
    case("html", f, "kommentartekst 333", found=False, why="a comment is not the page's")
    case("html", f, "titteltekst 444", found=False, why="an attribute is not the page's text")
    case("html", f, "feilutbetalingene til 4,2 prosent", found=False, why="a figure that is not there")


slides()
workbook()
texts()
html()
(OUT / "cases.json").write_text(json.dumps(cases, ensure_ascii=False, indent=1), encoding="utf-8")
print(len(cases), "cases in", OUT)
