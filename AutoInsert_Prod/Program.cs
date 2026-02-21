using ClosedXML.Excel;

const int SheetsPerFile = 5;
const string dailyDir = @"c:\Users\INC_ITServicePWRApps\Documents\EDI";
const string dailyFileSuffix = " Stats - converted.xlsx";
const string connSrc = @"c:\Users\INC_ITServicePWRApps\Documents\EDI\EDI_Daily_Stats_MASTER_TEST.xlsx";

static string[] GetDefaultSheetSelectors() => ["DELFOR Summary", "ORDERS Summary", "DESADV Summary", "INVOIC Summary", "ORDRSP Summary"]; // sheets 2-6

static string GetDefaultDailyPath()
{
	var yesterday = DateTime.Today.AddDays(-0);
	var fileName = $"{yesterday:yyyy_MM_dd}{dailyFileSuffix}";
	return Path.Combine(dailyDir, fileName);
}

static string GetUpdatedCopyPath(string inputPath)
{
	var directory = Path.GetDirectoryName(inputPath) ?? ".";
	var fileNameWithoutExt = Path.GetFileNameWithoutExtension(inputPath);
	var ext = Path.GetExtension(inputPath);
	var candidate = Path.Combine(directory, $"{fileNameWithoutExt}_updated{ext}");

	if (!File.Exists(candidate))
		return candidate;

	var stamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
	return Path.Combine(directory, $"{fileNameWithoutExt}_updated_{stamp}{ext}");
}

static int GetLastUsedRowNumber(IXLWorksheet worksheet)
{
	var used = worksheet.RangeUsed();
	return used?.RangeAddress.LastAddress.RowNumber ?? 0;
}

static (int firstCol, int lastCol, int lastRow) GetUsedBounds(IXLWorksheet worksheet)
{
	var used = worksheet.RangeUsed();
	if (used is null)
		return (0, 0, 0);

	return (
		used.RangeAddress.FirstAddress.ColumnNumber,
		used.RangeAddress.LastAddress.ColumnNumber,
		used.RangeAddress.LastAddress.RowNumber
	);
}

static int GetSourceRowNumberForSheet(string sheetName)
{
	// User requirement: DELFOR + ORDERS use row 17; the other three use row 7.
	if (sheetName.Contains("DELFOR", StringComparison.OrdinalIgnoreCase))
		return 17;
	if (sheetName.Contains("ORDERS", StringComparison.OrdinalIgnoreCase))
		return 17;

	return 7;
}

static bool ShouldSkipColumnC(string sheetName)
{
	return sheetName.Contains("DELFOR", StringComparison.OrdinalIgnoreCase)
		|| sheetName.Contains("ORDERS", StringComparison.OrdinalIgnoreCase);
}

static int AppendRowValues(IXLWorksheet source, IXLWorksheet destination, int sourceRow, bool skipColumnC)
{
	var (firstCol, lastCol, _) = GetUsedBounds(source);
	if (lastCol == 0)
		return 0;

	var scootLeft = skipColumnC && lastCol >= 3;
	var mappedCells = new List<(int destCol, XLCellValue value)>();
	var hasAnyValue = false;

	for (var c = firstCol; c <= lastCol; c++)
	{
		if (scootLeft && c == 3)
			continue;

		var destCol = scootLeft && c > 3 ? c - 1 : c;
		var srcCell = source.Cell(sourceRow, c);
		mappedCells.Add((destCol, srcCell.Value));
		if (!srcCell.IsEmpty())
			hasAnyValue = true;
	}

	if (!hasAnyValue)
		return 0;

	var destinationRow = GetLastUsedRowNumber(destination) + 1;
	foreach (var (destCol, value) in mappedCells)
		destination.Cell(destinationRow, destCol).Value = value;

	if (scootLeft)
	{
		// Clear column C if nothing mapped into it (e.g., if there was no column D to scoot over).
		if (!mappedCells.Any(x => x.destCol == 3))
			destination.Cell(destinationRow, 3).Clear(XLClearOptions.Contents);

		// Clear the last column since values have shifted left.
		destination.Cell(destinationRow, lastCol).Clear(XLClearOptions.Contents);
	}

	// // Add a small marker in the next column (if available) to show it was inserted.
	// destination.Cell(destinationRow, lastCol + 1).Value = $"Inserted {DateTime.Now:yyyy-MM-dd HH:mm:ss}";
	return 1;
}

static int Usage(string? error = null)
{
	if (!string.IsNullOrWhiteSpace(error))
	{
		Console.Error.WriteLine(error);
		Console.Error.WriteLine();
	}

	Console.WriteLine("Notes:");
	Console.WriteLine("  - <sheetX> can be a sheet name (e.g. Sheet1) or a 1-based index (e.g. 2). ");
	Console.WriteLine($"  - If you omit sheet selectors, these are used (sheets 2-6): {string.Join(", ", GetDefaultSheetSelectors())}");
	Console.WriteLine($"  - With no args, defaults to (prior day):\n      A: {GetDefaultDailyPath()}\n      B: {connSrc}");
	return 2;
}

static IXLWorksheet GetWorksheet(XLWorkbook workbook, string? selector)
{
	if (string.IsNullOrWhiteSpace(selector))
		return workbook.Worksheets.First();

	if (int.TryParse(selector, out var index))
	{
		if (index < 1 || index > workbook.Worksheets.Count)
			throw new ArgumentOutOfRangeException(nameof(selector), $"Worksheet index {index} is out of range (1..{workbook.Worksheets.Count}).");

		return workbook.Worksheet(index);
	}

	var byName = workbook.Worksheets.FirstOrDefault(w => string.Equals(w.Name, selector, StringComparison.OrdinalIgnoreCase));
	if (byName is null)
		throw new ArgumentException($"Worksheet not found: '{selector}'", nameof(selector));

	return byName;
}

static IReadOnlyList<IXLWorksheet> GetFiveWorksheets(XLWorkbook workbook, IReadOnlyList<string>? selectors)
{
	if (selectors is { Count: > 0 })
	{
		if (selectors.Count != SheetsPerFile)
			throw new ArgumentException($"Expected exactly {SheetsPerFile} sheet selectors, got {selectors.Count}.");

		return selectors.Select(s => GetWorksheet(workbook, s)).ToArray();
	}

	var defaultSelectors = GetDefaultSheetSelectors();
	if (defaultSelectors.Length != SheetsPerFile)
		throw new InvalidOperationException($"Default sheet selector list must contain exactly {SheetsPerFile} items.");

	return defaultSelectors.Select(s => GetWorksheet(workbook, s)).ToArray();
}

static void PrintPreview(string label, IXLWorksheet worksheet, int maxRows = 10, int maxCols = 10)
{
	var used = worksheet.RangeUsed();
	if (used is null)
	{
		Console.WriteLine($"{label}: '{worksheet.Name}' is empty.");
		return;
	}

	var firstRow = used.RangeAddress.FirstAddress.RowNumber;
	var firstCol = used.RangeAddress.FirstAddress.ColumnNumber;
	var lastRow = used.RangeAddress.LastAddress.RowNumber;
	var lastCol = used.RangeAddress.LastAddress.ColumnNumber;

	var rowCount = lastRow - firstRow + 1;
	var colCount = lastCol - firstCol + 1;

	Console.WriteLine($"{label}: '{worksheet.Name}' used range = {rowCount} rows x {colCount} cols");

	var previewLastRow = Math.Min(lastRow, firstRow + maxRows - 1);
	var previewLastCol = Math.Min(lastCol, firstCol + maxCols - 1);

	for (var r = firstRow; r <= previewLastRow; r++)
	{
		var values = new List<string>(capacity: previewLastCol - firstCol + 1);
		for (var c = firstCol; c <= previewLastCol; c++)
		{
			var cellText = worksheet.Cell(r, c).GetFormattedString();
			values.Add(cellText);
		}
		Console.WriteLine(string.Join("\t", values));
	}
}

try
{
	var path1 = GetDefaultDailyPath();
	var path2 = connSrc;
	IReadOnlyList<string>? selectors1 = null;
	IReadOnlyList<string>? selectors2 = null;

	if (args.Length != 0)
	{
		if (args.Length is not (2 or 12))
			return Usage("Expected 0 args (use hardcoded paths), 2 args (two files), or 12 args (two files + 5 sheet selectors each).");

		path1 = args[0];
		path2 = args[1];

		if (args.Length == 12)
		{
			selectors1 = args.Skip(1).Take(SheetsPerFile).ToArray();
			path2 = args[1 + SheetsPerFile];
			selectors2 = args.Skip(2 + SheetsPerFile).Take(SheetsPerFile).ToArray();
		}
	}

	if (!File.Exists(path1))
		return Usage($"File not found: {path1}");
	if (!File.Exists(path2))
		return Usage($"File not found: {path2}");

	using var wb1 = new XLWorkbook(path1);
	using var wb2 = new XLWorkbook(path2);

	var wsList1 = GetFiveWorksheets(wb1, selectors1);
	var wsList2 = GetFiveWorksheets(wb2, selectors2);

	for (var i = 0; i < wsList1.Count; i++)
	{
		PrintPreview($"File A - {wsList1[i].Name}", wsList1[i]);
		Console.WriteLine();
	}

	for (var i = 0; i < wsList2.Count; i++)
	{
		PrintPreview($"File B - {wsList2[i].Name}", wsList2[i]);
		if (i < wsList2.Count - 1)
			Console.WriteLine();
	}

	Console.WriteLine();
	Console.WriteLine("Appending 1 new row into each connected sheet...");

	var insertedTotal = 0;
	for (var i = 0; i < SheetsPerFile; i++)
	{
		var sourceRow = GetSourceRowNumberForSheet(wsList1[i].Name);
		var skipColumnC = ShouldSkipColumnC(wsList1[i].Name);
		var inserted = AppendRowValues(wsList1[i], wsList2[i], sourceRow, skipColumnC);
		insertedTotal += inserted;
		var note = skipColumnC ? ", col C skipped + scooted left" : string.Empty;
		Console.WriteLine($"  {wsList2[i].Name}: {(inserted == 1 ? $"inserted (from row {sourceRow}{note})" : $"skipped (row {sourceRow} empty)")}");
	}

	var outputConnPath = GetUpdatedCopyPath(path2);
	wb2.SaveAs(outputConnPath);
	Console.WriteLine($"Saved updated connected workbook: {outputConnPath}");
	Console.WriteLine($"Rows inserted: {insertedTotal}/{SheetsPerFile}");

	return 0;
}
catch (Exception ex)
{
	Console.Error.WriteLine("Failed to read Excel.");
	Console.Error.WriteLine(ex.Message);
	return 1;
}
