---
name: cite
description: Show where each claim comes from, as links that open the source at the cited words. Use when the user asks where something is said, asks for sources or citations, asks you to show your work or back up an answer, or says "vis kildene", "hvor står det", "show me where it says that", "cite your sources". Covers Word, PowerPoint and Excel files, PDFs, text and code files, HTML files and web pages. Windows only.
---

# Cite with Show your work

The user wants to check what you said against the source. Give them links
that open each source at the exact words.

## What to do

1. Decide which claims to back: the ones in your last answer, or the ones the
   user points at. Take the claims that matter — figures, findings, dates,
   anything quoted — not every sentence.
2. For each claim, call the `show_source` tool
   (`mcp__show-your-work__show_source`) with the source's absolute path and
   the source's own words, verbatim. Make the calls in parallel.
   - Excel: give `cell` (`Sheet!B2`, a row `Sheet!5:5`, a column `Sheet!C:C`,
     a block, or a sheet alone `Sheet!`) and, as `quote`, the value you say is
     there.
   - PDF: give `page`.
   - If you do not have the exact words, read the passage again first. Do not
     quote from memory.
3. Write the answer again, with each number the tool returned, like `[3]`,
   right after the claim it backs. The user does not see the number: they see
   a link reading the place (`p. 7`, `row 250`, `line 19`).

## What the tool's answers mean

- **Found word for word**: the words are in the source at that place.
- **NOT FOUND**: the words are not in the source. Correct the claim or tell
  the user it is unverified. Never keep a claim the source does not carry.
- **Noted, NOT checked**: a web page is never fetched, so nothing was checked.
  Say so if the claim matters.
- **Could not be checked**: the program did not answer, or the file is
  missing. Tell the user which source, and that the claim is unverified.

## If the tool is not there

The plugin offers nothing when the user has set **Cite** to **Off** in the
Sources pane, and on systems other than Windows. Then cite in plain text —
file, page or cell, and the quoted words — and tell the user that the links
are off and that `/sources` opens the pane where they are turned on.
