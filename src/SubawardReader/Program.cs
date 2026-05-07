using SubawardReader;

string folderPath = args.Length > 0
    ? args[0]
    : Path.Combine(AppContext.BaseDirectory, "Samples");

if (!Directory.Exists(folderPath))
{
    Console.WriteLine($"Folder not found: {folderPath}");
    Console.WriteLine("Usage: dotnet run --project src/SubawardReader -- <folderPath>");
    return;
}

var excelFiles = Directory
    .EnumerateFiles(folderPath, "*.xlsx", SearchOption.TopDirectoryOnly)
    .Where(path => !Path.GetFileName(path).StartsWith("~$", StringComparison.Ordinal))
    .OrderBy(Path.GetFileName)
    .ToList();

if (excelFiles.Count == 0)
{
    Console.WriteLine($"No .xlsx files found in: {folderPath}");
    return;
}

var parser = new SubawardParser();
var allRecords = new List<SubawardRecord>();

Console.WriteLine("SUBAWARD READER");
Console.WriteLine(new string('=', 72));
Console.WriteLine();

foreach (string filePath in excelFiles)
{
    var records = parser.ParseFile(filePath).ToList();
    allRecords.AddRange(records);

    Console.WriteLine($"File: {Path.GetFileName(filePath)}");
    Console.WriteLine(new string('-', 72));

    if (records.Count == 0)
    {
        Console.WriteLine("  No subrecipients found.");
    }
    else
    {
        foreach (var record in records)
        {
            Console.WriteLine($"  - {record.SubrecipientName,-30} {record.Amount,15:C}");
        }
    }

    Console.WriteLine();
}

Console.WriteLine("DISTINCT SUBRECIPIENT TOTALS ACROSS ALL FILES");
Console.WriteLine(new string('=', 72));

var totals = allRecords
    .GroupBy(r => r.SubrecipientName, StringComparer.OrdinalIgnoreCase)
    .Select(g => new
    {
        Name = g.First().SubrecipientName,
        Total = g.Sum(x => x.Amount)
    })
    .OrderBy(x => x.Name)
    .ToList();

if (totals.Count == 0)
{
    Console.WriteLine("No subaward totals to display.");
}
else
{
    Console.WriteLine($"{"Subrecipient",-35} {"Total Amount",15}");
    Console.WriteLine(new string('-', 72));

    foreach (var total in totals)
    {
        Console.WriteLine($"{total.Name,-35} {total.Total,15:C}");
    }
}
