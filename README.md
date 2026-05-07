# Subaward Reader

Small .NET 9 console application that reads budget spreadsheets from a folder, prints each spreadsheet's subrecipients, and then prints distinct subrecipient totals across all spreadsheets.

---

# Tech Stack

- .NET 9
- C#
- ClosedXML for reading Excel workbooks
- xUnit for unit testing

---

# Repository

GitHub Repository:

```text
https://github.com/omkarvchalke/Subaward-Reader-Submission
```

---

# How to Run

Clone the repository:

```bash
git clone https://github.com/omkarvchalke/Subaward-Reader-Submission.git
cd Subaward-Reader-Submission
```

Restore packages and build:

```bash
dotnet restore
dotnet build
```

Run the application using the provided sample spreadsheets:

```bash
dotnet run --project src/SubawardReader -- Samples
```

You can also pass any folder containing spreadsheets in the same format:

```bash
dotnet run --project src/SubawardReader -- /path/to/spreadsheet-folder
```

---

# How to Run Tests

```bash
dotnet test
```

The included unit test confirms that `SubawardBudgetExample1.xlsx` contains the four expected subrecipients:

- Indiana
- Mayo
- Purdue
- Florida

---

# Expected Console Output Format

```text
SUBAWARD READER
========================================================================

File: SubawardBudgetExample1.xlsx
------------------------------------------------------------------------
  - Indiana                              $126,904.00
  - Mayo                                  $20,637.00
  - Purdue                                $64,101.00
  - Florida                               $74,423.00

DISTINCT SUBRECIPIENT TOTALS ACROSS ALL FILES
========================================================================
Subrecipient                           Total Amount
------------------------------------------------------------------------
Florida                                  $74,423.00
Indiana                                 $126,904.00
Mayo                                    $136,736.00
Purdue                                   $89,101.00
```

---

# Parsing Approach

The parser does not hardcode exact row numbers. It:

1. Finds the `Other Direct Costs` section.
2. Scans rows inside that section.
3. Finds rows whose text starts with `Subaward`.
4. Extracts the subrecipient name either:
   - from the same cell after `Subaward:`
   - or from the next non-empty text cell to the right.
5. Calculates the subaward total by using:
   - the rightmost numeric amount on the `Subaward` row
   - plus the rightmost numeric amount on the immediately following `Exempt Subaward Costs (>$25k)` row when that row exists.

---

# Assumptions

- Input files are `.xlsx` files.
- The relevant section is labeled `Other Direct Costs`.
- Subaward rows begin with `Subaward` or `Subaward:`.
- Subrecipient names may appear either:
  - in the same cell as `Subaward:`
  - or in the next text cell to the right.
- The rightmost numeric value on a row represents that row's total amount.
- The full subaward received by a subrecipient includes:
  - the `Subaward` row
  - and the following `Exempt Subaward Costs (>$25k)` row when present.
- Blank subaward-name rows are ignored.
- Temporary Excel lock files beginning with `~$` are ignored.

---

# Questions I Would Ask in Real Work

1. Should the final total include only the amount on the `Subaward:` row, or should it also include the following `Exempt Subaward Costs (>$25k)` row?

2. Is the rightmost numeric column always the final total, or can future spreadsheets label totals differently?

3. Should blank `Subaward:` rows be treated as data-quality errors instead of being ignored?

4. Can workbooks contain multiple worksheets with valid budget data, and should all worksheets always be scanned?

5. Should the console output also support CSV or Excel export for non-technical staff?

---

# Design Decisions

## Why ClosedXML?

ClosedXML provides:

- Simple Excel workbook handling
- Readable APIs
- Reliable `.xlsx` support
- Easy worksheet traversal

## Why Dynamic Parsing?

The requirements specified support for:

```text
Variable number of subaward rows
```

Because of this, the parser dynamically scans spreadsheet sections instead of hardcoding row indexes.

---

# Author

Omkar Vilas Chalke
