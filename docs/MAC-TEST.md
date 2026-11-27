# Testing Show your work on a Mac

The Mac version was written on Windows. What reads the files, and showing
text, HTML and web pages, is checked on a Mac; what drives Word, Excel and
PowerPoint has never run against the programs (that takes a Mac with a
licensed Office). This is how to check it, in about fifteen minutes. Nothing
is installed, and nothing leaves the machine.

You need a Mac with Word, Excel and PowerPoint. VS Code and
[PDF Scholar](https://github.com/emilmsh/pdf-scholar) are optional.

## 1. Get the files

```
git clone https://github.com/emilmsh/show-your-work.git
cd show-your-work
```

## 2. Look every case up (nothing is shown)

```
sh scripts/test/check-mac.sh > result.txt
```

The first run asks, once per program, whether Terminal may control Word, Excel
and PowerPoint. Answer **OK**. If you answered no by mistake: System Settings,
**Privacy & Security**, **Automation**, and switch them on under Terminal.

Word and Excel may come up for a moment and close again. Note whether they did,
and whether anything was left open or asked you something.

## 3. Open one source of each kind at its words

```
sh scripts/test/check-mac.sh show > shown.txt
```

Seven sources open, four seconds apart. For each, note what you saw:

| # | Source | What should happen |
| --- | --- | --- |
| 1 | `report.docx` | Word, page 1, the sentence «Attendance rose by 4.2 percentage points in the pilot schools» selected |
| 2 | `model.xlsx` | Excel, sheet Results, cell B2 selected |
| 3 | `model.xlsx` | Excel, sheet Results, column B selected, with the cell holding 4.2 active |
| 4 | `slides.pptx` | PowerPoint on slide 4 (no words selected: PowerPoint on a Mac cannot) |
| 5 | `analysis.R` | VS Code at line 10 if it is installed, else the file in the default text editor |
| 6 | `notes.md` | VS Code at line 12, or the default text editor |
| 7 | example.com | The default browser, with «Example Domain» marked on the page |

## 4. Send back

`result.txt`, `shown.txt` and your notes on the seven. A `FAIL` line, an
`error` or a `place` in the output is exactly what is needed: each says which
call a Mac refused.

## 5. In Claude Code

Install the plugin as the README says (version 0.4.0 or later). Then, in the
`demo` folder, ask the question from the README and press the links in the
answer. The first press for each program asks whether Claude Code (or your
terminal) may control it: answer **OK**.
