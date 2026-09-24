# Reader test fixtures

Small workbooks, each pinning down one thing `read_excel_all()` has to get
right. They are committed rather than generated at test time so the tests do
not depend on the library being tested to build its own inputs.

Most were hand-written as raw OOXML and zipped, precisely so they contain
structures Excel produces but a writing library generally will not (a cell
element with no value, a shared-formula slave with no formula text, a
number format applied to serial zero). That is deliberate: these are the
shapes that broke things.

| file | what it holds | why |
|---|---|---|
| `comment-on-blank-cell.xlsx` | `<c r="F4" s="2"/>` — a styled cell with no value — plus a comment on F4 | The bug that started all this. `tidyxl::xlsx_cells(include_blank_cells = FALSE)` miscounts and overruns its output vectors, failing with `attempt to set index N/N in SET_STRING_ELT`. Written by openxlsx2, so this is a shape real tools emit. |
| `offset-origin.xlsx` | Data starting at D5; rows 1–4 and columns A–C entirely absent | The grid must stay anchored at A1. `readxl::read_excel()` and `openxlsx::read.xlsx()` both trim these away, which is why neither can be used here. |
| `value-types.xlsx` | Float that must not be rounded (`0.30000000000000004`), scientific notation (`1e-9`, `0.0000123`), a 16-digit integer, leading-zero text (`00123`), a datetime, a `#DIV/0!` error, a boolean, a percentage-formatted number | Cell values are rendered to character. Every one of these is a way that can go wrong silently. |
| `shared-formula.xlsx` | B1 carries `A1*2` as a shared-formula master; B2 and B3 inherit it via `si="0"` and carry no formula text | Characterises the one known behavioural gap: tidyxl reconstructs the inherited formulas and shifts their references, openxlsx2 does not. |
| `shared-formula-empty-result.xlsx` | The same, but the formula's cached result is the empty string | This is the shape of the tracker's error-check columns, and the source of the `""` vs `NA` differences seen across the board files. |
| `date-serial-edges.xlsx` | Serials 0, 0.25, 0.5, 0.75 and 1 under time (`h:mm`, `h:mm:ss`) and date number formats, plus one ordinary datetime | Serial 0 is not a real date. It is the only value the two backends ever disagreed about across 7.4 million production cells. |
