# Subaward Reader

Small .NET 9 console application that reads budget spreadsheets from a folder, prints each spreadsheet's subrecipients, and then prints distinct subrecipient totals across all spreadsheets.

## Tech stack

- .NET 9
- C#
- ClosedXML for reading Excel workbooks
- xUnit for unit testing

## How to run

```bash
git clone <your-public-repository-url>
cd SubawardReader

dotnet restore
dotnet build

dotnet run --project src/SubawardReader -- Samples
```

You can also pass any folder that contains spreadsheets in the same format:

```bash
dotnet run --project src/SubawardReader -- /path/to/spreadsheet-folder
```

## How to run tests

```bash
dotnet test
```

The included unit test confirms that `SubawardBudgetExample1.xlsx` contains the four expected subrecipients: Indiana, Mayo, Purdue, and Florida.

## Expected console output format

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

## Parsing approach

The parser does not hardcode exact row numbers. It:

1. Finds the `Other Direct Costs` section.
2. Scans rows inside that section.
3. Finds rows whose text starts with `Subaward`.
4. Extracts the subrecipient name either from the same cell after `Subaward:` or from the next non-empty text cell to the right.
5. Calculates the subaward total by using the rightmost numeric amount on the `Subaward` row plus the rightmost numeric amount on the immediately following `Exempt Subaward Costs (>$25k)` row when that row exists.

## Assumptions

- Input files are `.xlsx` files.
- The relevant section is labeled `Other Direct Costs`.
- Subaward rows begin with `Subaward` or `Subaward:`.
- Subrecipient names may appear either in the same cell as `Subaward:` or in the next text cell to the right.
- The rightmost numeric value on a row represents that row's total amount. This avoids double-counting spreadsheets that show year-by-year values plus a final total column.
- The full subaward received by a subrecipient includes the `Subaward` row and the following `Exempt Subaward Costs (>$25k)` row, when present.
- Blank subaward-name rows are ignored.
- Temporary Excel lock files beginning with `~$` are ignored.

## Questions I would ask in real work

1. Should the final total include only the amount on the `Subaward:` row, or should it include the following `Exempt Subaward Costs (>$25k)` row as part of the recipient's total subaward amount?
2. Is the rightmost numeric column always the final total, or can the total column be labeled differently across future spreadsheets?
3. Should blank `Subaward:` rows be treated as data quality errors, or should they be ignored as this implementation does?
4. Can there be multiple worksheets per workbook with valid budget data, and should all worksheets always be scanned?
5. Should the console output also be exportable to CSV or Excel for non-technical staff?
