using ClosedXML.Excel;

namespace SubawardReader;

public sealed class SubawardParser
{
    public IEnumerable<SubawardRecord> ParseFile(string filePath)
    {
        if (string.IsNullOrWhiteSpace(filePath))
        {
            throw new ArgumentException("File path cannot be empty.", nameof(filePath));
        }

        if (!File.Exists(filePath))
        {
            throw new FileNotFoundException("Spreadsheet file was not found.", filePath);
        }

        using var workbook = new XLWorkbook(filePath);

        foreach (var worksheet in workbook.Worksheets)
        {
            foreach (var record in ParseWorksheet(filePath, worksheet))
            {
                yield return record;
            }
        }
    }

    private static IEnumerable<SubawardRecord> ParseWorksheet(string filePath, IXLWorksheet worksheet)
    {
        var usedRange = worksheet.RangeUsed();
        if (usedRange is null)
        {
            yield break;
        }

        int firstRow = usedRange.FirstRowUsed().RowNumber();
        int lastRow = usedRange.LastRowUsed().RowNumber();
        int firstColumn = usedRange.FirstColumnUsed().ColumnNumber();
        int lastColumn = usedRange.LastColumnUsed().ColumnNumber();

        int sectionStartRow = FindOtherDirectCostsStartRow(worksheet, firstRow, lastRow, firstColumn, lastColumn);
        if (sectionStartRow == -1)
        {
            yield break;
        }

        for (int row = sectionStartRow + 1; row <= lastRow; row++)
        {
            string rowText = GetRowText(worksheet, row, firstColumn, lastColumn);

            if (IsEndOfOtherDirectCostsSection(rowText))
            {
                yield break;
            }

            for (int column = firstColumn; column <= lastColumn; column++)
            {
                string cellText = GetCellText(worksheet.Cell(row, column));

                if (!cellText.StartsWith("Subaward", StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                string? name = ExtractSubrecipientName(worksheet, row, column, lastColumn);
                if (string.IsNullOrWhiteSpace(name))
                {
                    continue;
                }

                decimal amount = GetSubawardTotalAmount(worksheet, row, firstColumn, lastColumn);
                yield return new SubawardRecord(Path.GetFileName(filePath), name, amount);
                break;
            }
        }
    }

    private static int FindOtherDirectCostsStartRow(
        IXLWorksheet worksheet,
        int firstRow,
        int lastRow,
        int firstColumn,
        int lastColumn)
    {
        for (int row = firstRow; row <= lastRow; row++)
        {
            string rowText = GetRowText(worksheet, row, firstColumn, lastColumn);

            if (rowText.Contains("Other Direct Costs", StringComparison.OrdinalIgnoreCase))
            {
                return row;
            }
        }

        return -1;
    }

    private static bool IsEndOfOtherDirectCostsSection(string rowText)
    {
        if (string.IsNullOrWhiteSpace(rowText))
        {
            return false;
        }

        return rowText.Contains("Total Direct Costs", StringComparison.OrdinalIgnoreCase)
            || rowText.StartsWith("H.", StringComparison.OrdinalIgnoreCase)
            || rowText.StartsWith("I.", StringComparison.OrdinalIgnoreCase)
            || rowText.StartsWith("J.", StringComparison.OrdinalIgnoreCase)
            || rowText.StartsWith("K.", StringComparison.OrdinalIgnoreCase);
    }

    private static string? ExtractSubrecipientName(
        IXLWorksheet worksheet,
        int row,
        int subawardColumn,
        int lastColumn)
    {
        string cellText = GetCellText(worksheet.Cell(row, subawardColumn));
        int colonIndex = cellText.IndexOf(':');

        if (colonIndex >= 0)
        {
            string nameAfterColon = cellText[(colonIndex + 1)..].Trim();
            if (!string.IsNullOrWhiteSpace(nameAfterColon))
            {
                return CleanName(nameAfterColon);
            }
        }

        for (int column = subawardColumn + 1; column <= lastColumn; column++)
        {
            string possibleName = GetCellText(worksheet.Cell(row, column));
            if (!string.IsNullOrWhiteSpace(possibleName) && !IsNumericCell(worksheet.Cell(row, column)))
            {
                return CleanName(possibleName);
            }
        }

        return null;
    }

    private static decimal GetSubawardTotalAmount(
        IXLWorksheet worksheet,
        int subawardRow,
        int firstColumn,
        int lastColumn)
    {
        decimal subawardLineAmount = GetRightMostNumericAmount(worksheet, subawardRow, firstColumn, lastColumn);

        int nextRow = subawardRow + 1;
        string nextRowText = GetRowText(worksheet, nextRow, firstColumn, lastColumn);

        decimal exemptAmount = nextRowText.Contains("Exempt Subaward Costs", StringComparison.OrdinalIgnoreCase)
            ? GetRightMostNumericAmount(worksheet, nextRow, firstColumn, lastColumn)
            : 0m;

        return subawardLineAmount + exemptAmount;
    }

    private static decimal GetRightMostNumericAmount(
        IXLWorksheet worksheet,
        int row,
        int firstColumn,
        int lastColumn)
    {
        for (int column = lastColumn; column >= firstColumn; column--)
        {
            var cell = worksheet.Cell(row, column);

            if (TryGetDecimal(cell, out decimal amount))
            {
                return amount;
            }
        }

        return 0m;
    }

    private static bool TryGetDecimal(IXLCell cell, out decimal value)
    {
        if (cell.TryGetValue(out double numericValue))
        {
            value = Convert.ToDecimal(numericValue);
            return true;
        }

        string text = GetCellText(cell)
            .Replace("$", string.Empty, StringComparison.Ordinal)
            .Replace(",", string.Empty, StringComparison.Ordinal)
            .Trim();

        return decimal.TryParse(text, out value);
    }

    private static bool IsNumericCell(IXLCell cell)
    {
        return TryGetDecimal(cell, out _);
    }

    private static string GetRowText(IXLWorksheet worksheet, int row, int firstColumn, int lastColumn)
    {
        var values = new List<string>();

        for (int column = firstColumn; column <= lastColumn; column++)
        {
            string text = GetCellText(worksheet.Cell(row, column));
            if (!string.IsNullOrWhiteSpace(text))
            {
                values.Add(text);
            }
        }

        return string.Join(" ", values).Trim();
    }

    private static string GetCellText(IXLCell cell)
    {
        return cell.GetFormattedString().Trim();
    }

    private static string CleanName(string name)
    {
        return name.Trim().Trim(':').Trim();
    }
}
