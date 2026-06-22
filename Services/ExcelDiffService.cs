using System.Globalization;
using System.Security.Cryptography;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using ExcelCompare.Models;
using Microsoft.AspNetCore.Components.Forms;

namespace ExcelCompare.Services;

public sealed class ExcelDiffService
{
    public const long MaxFileSize = 10 * 1024 * 1024;

    private const int ShiftLookAhead = 24;
    private const int MaxPerShiftBlock = 250;
    private const int MaxMacroLineDiffsPerModule = 200;

    public async Task<FileMetadata> ReadMetadataAsync(IBrowserFile file)
    {
        ValidateFile(file);

        await using var browserStream = file.OpenReadStream(MaxFileSize);
        using var memory = new MemoryStream((int)Math.Min(file.Size, MaxFileSize));
        await browserStream.CopyToAsync(memory);
        memory.Position = 0;

        using var document = SpreadsheetDocument.Open(memory, false);
        var workbookPart = document.WorkbookPart
            ?? throw new InvalidOperationException("Excelブックの構造を読み取れませんでした。");

        var workbook = workbookPart.Workbook
            ?? throw new InvalidOperationException("Excelブックの構造を読み取れませんでした。");
        var sheets = workbook.Sheets?.Elements<Sheet>().ToList() ?? [];
        return new FileMetadata(
            file.Name,
            file.Size,
            Path.GetExtension(file.Name).ToLowerInvariant(),
            sheets.Count,
            sheets.Select(sheet => sheet.Name?.Value ?? "(名称なし)").ToList(),
            workbookPart.VbaProjectPart is not null);
    }

    public async Task<CompareResult> CompareAsync(
        IBrowserFile sourceFile,
        IBrowserFile targetFile,
        ComparisonOptions options,
        IProgress<CompareProgress>? progress = null)
    {
        if (!options.IncludeData &&
            !options.IncludeFormatting &&
            !options.IncludeStructure &&
            !options.IncludeMacros)
        {
            throw new InvalidOperationException("比較対象を1つ以上選択してください。");
        }

        progress?.Report(new("Excel読込中...", 8));
        var source = await ReadWorkbookAsync(sourceFile, includeMacroSource: options.IncludeMacros);
        progress?.Report(new("比較元の解析完了", 25));
        await Task.Yield();

        var target = await ReadWorkbookAsync(targetFile, includeMacroSource: options.IncludeMacros);
        progress?.Report(new("シート解析中...", 45));
        await Task.Yield();

        var records = new List<DiffRecord>(capacity: 256);

        if (options.IncludeStructure)
        {
            CompareStructure(source, target, records);
        }

        if (options.IncludeData)
        {
            CompareDefinedNames(source, target, records);
        }

        if (options.IncludeData || options.IncludeFormatting)
        {
            progress?.Report(new("差分比較中...", 65));
            CompareSheets(source, target, options, records);
        }

        if (options.IncludeMacros)
        {
            progress?.Report(new("マクロ比較中...", 82));
            CompareMacros(source, target, records);
        }

        progress?.Report(new("結果生成中...", 94));
        var summary = new DiffSummary(
            records.Count(row => row.Category == DiffCategory.Structure),
            records.Count(row => row.Category == DiffCategory.Cell),
            records.Count(row => row.Category == DiffCategory.Formatting),
            records.Count(row => row.Category == DiffCategory.Macro),
            records.Count(row => row.Category == DiffCategory.Metadata));

        progress?.Report(new("完了", 100));
        return new CompareResult(source.Metadata, target.Metadata, summary, records);
    }

    private static async Task<WorkbookSnapshot> ReadWorkbookAsync(IBrowserFile file, bool includeMacroSource)
    {
        ValidateFile(file);

        await using var browserStream = file.OpenReadStream(MaxFileSize);
        using var memory = new MemoryStream((int)Math.Min(file.Size, MaxFileSize));
        await browserStream.CopyToAsync(memory);
        memory.Position = 0;

        using var document = SpreadsheetDocument.Open(memory, false);
        var workbookPart = document.WorkbookPart
            ?? throw new InvalidOperationException("Excelブックの構造を読み取れませんでした。");

        var workbook = workbookPart.Workbook
            ?? throw new InvalidOperationException("Excelブックの構造を読み取れませんでした。");
        var sheets = workbook.Sheets?.Elements<Sheet>().ToList() ?? [];
        var sheetSnapshots = new List<SheetSnapshot>(sheets.Count);
        var sharedStrings = LoadSharedStrings(workbookPart);
        var styleLookup = BuildStyleLookup(workbookPart);

        for (var index = 0; index < sheets.Count; index++)
        {
            var sheet = sheets[index];
            if (sheet.Id?.Value is not { Length: > 0 } relationshipId ||
                workbookPart.GetPartById(relationshipId) is not WorksheetPart worksheetPart)
            {
                continue;
            }

            sheetSnapshots.Add(ReadSheet(sheet, worksheetPart, index + 1, sharedStrings, styleLookup));
        }

        var metadata = new FileMetadata(
            file.Name,
            file.Size,
            Path.GetExtension(file.Name).ToLowerInvariant(),
            sheetSnapshots.Count,
            sheetSnapshots.Select(sheet => sheet.Name).ToList(),
            workbookPart.VbaProjectPart is not null);

        return new WorkbookSnapshot(
            metadata,
            sheetSnapshots,
            ReadDefinedNames(workbookPart),
            ReadMacro(workbookPart, includeMacroSource));
    }

    private static void ValidateFile(IBrowserFile file)
    {
        var extension = Path.GetExtension(file.Name);
        if (!extension.Equals(".xlsx", StringComparison.OrdinalIgnoreCase) &&
            !extension.Equals(".xlsm", StringComparison.OrdinalIgnoreCase))
        {
            throw new InvalidOperationException(".xlsx または .xlsm 形式のみ対応しています。");
        }

        if (file.Size > MaxFileSize)
        {
            throw new InvalidOperationException("1ファイルあたり10MB以下のExcelファイルを選択してください。");
        }
    }

    private static IReadOnlyList<string> LoadSharedStrings(WorkbookPart workbookPart)
    {
        var table = workbookPart.SharedStringTablePart?.SharedStringTable;
        return table is null
            ? []
            : table.Elements<SharedStringItem>().Select(item => item.InnerText ?? string.Empty).ToList();
    }

    private static SheetSnapshot ReadSheet(
        Sheet sheet,
        WorksheetPart worksheetPart,
        int position,
        IReadOnlyList<string> sharedStrings,
        IReadOnlyDictionary<uint, string> styleLookup)
    {
        var cellsByAddress = new Dictionary<string, CellSnapshot>(StringComparer.OrdinalIgnoreCase);
        var mutableRows = new Dictionary<uint, Dictionary<uint, CellSnapshot>>();

        var worksheet = worksheetPart.Worksheet
            ?? throw new InvalidOperationException("ワークシートの構造を読み取れませんでした。");

        foreach (var cell in worksheet.Descendants<Cell>())
        {
            var address = cell.CellReference?.Value;
            if (string.IsNullOrWhiteSpace(address))
            {
                continue;
            }

            var (column, row) = ParseCellReference(address);
            if (row == 0 || column == 0)
            {
                continue;
            }

            var snapshot = new CellSnapshot(
                address,
                row,
                column,
                ReadCellValue(cell, sharedStrings),
                cell.CellFormula?.Text ?? string.Empty,
                ReadStyle(cell, styleLookup));

            if (string.IsNullOrEmpty(snapshot.Value) &&
                string.IsNullOrEmpty(snapshot.Formula) &&
                string.IsNullOrEmpty(snapshot.StyleSignature))
            {
                continue;
            }

            cellsByAddress[address] = snapshot;
            if (!mutableRows.TryGetValue(row, out var rowCells))
            {
                rowCells = [];
                mutableRows[row] = rowCells;
            }

            rowCells[column] = snapshot;
        }

        var rows = mutableRows.ToDictionary(
            pair => pair.Key,
            pair => (IReadOnlyDictionary<uint, CellSnapshot>)pair.Value,
            EqualityComparer<uint>.Default);

        var rowSignatures = rows.ToDictionary(
            pair => pair.Key,
            pair => BuildRowSignature(pair.Value),
            EqualityComparer<uint>.Default);

        var columnSignatures = cellsByAddress.Values
            .GroupBy(cell => cell.Column)
            .ToDictionary(
                group => group.Key,
                group => BuildColumnSignature(group),
                EqualityComparer<uint>.Default);

        return new SheetSnapshot(
            sheet.Name?.Value ?? "(名称なし)",
            position,
            sheet.State?.Value.ToString() ?? "Visible",
            cellsByAddress,
            rowSignatures,
            columnSignatures,
            rows,
            ReadHiddenRows(worksheet),
            ReadHiddenColumns(worksheet),
            ReadDataValidations(worksheet),
            worksheet.Elements<SheetProtection>().FirstOrDefault()?.OuterXml ?? string.Empty);
    }

    private static string ReadCellValue(Cell cell, IReadOnlyList<string> sharedStrings)
    {
        if (cell.DataType?.Value == CellValues.SharedString &&
            int.TryParse(cell.CellValue?.Text, NumberStyles.Integer, CultureInfo.InvariantCulture, out var index) &&
            index >= 0 &&
            index < sharedStrings.Count)
        {
            return sharedStrings[index];
        }

        if (cell.DataType?.Value == CellValues.InlineString)
        {
            return cell.InlineString?.InnerText ?? string.Empty;
        }

        if (cell.DataType?.Value == CellValues.Boolean)
        {
            return cell.CellValue?.Text == "1" ? "TRUE" : "FALSE";
        }

        return cell.CellValue?.Text ?? string.Empty;
    }

    private static IReadOnlyDictionary<uint, string> BuildStyleLookup(WorkbookPart workbookPart)
    {
        var stylesheet = workbookPart.WorkbookStylesPart?.Stylesheet;
        var formats = stylesheet?.CellFormats?.Elements<CellFormat>().ToList() ?? [];
        var fills = stylesheet?.Fills?.Elements<Fill>().ToList() ?? [];
        var fonts = stylesheet?.Fonts?.Elements<Font>().ToList() ?? [];
        var customNumberFormats = stylesheet?.NumberingFormats?.Elements<NumberingFormat>()
            .Where(format => format.NumberFormatId?.Value is not null)
            .ToDictionary(format => format.NumberFormatId!.Value, format => format.FormatCode?.Value ?? string.Empty)
            ?? [];

        var result = new Dictionary<uint, string>();
        for (var index = 0; index < formats.Count; index++)
        {
            var format = formats[index];
            var fill = GetAt(fills, format.FillId?.Value);
            var font = GetAt(fonts, format.FontId?.Value);
            var numberFormatId = format.NumberFormatId?.Value ?? 0U;
            var numberFormat = customNumberFormats.GetValueOrDefault(numberFormatId, $"builtin:{numberFormatId}");

            var background = ColorKey(fill?.PatternFill?.ForegroundColor) ??
                ColorKey(fill?.PatternFill?.BackgroundColor) ??
                string.Empty;
            var textColor = ColorKey(font?.Color) ?? string.Empty;

            result[(uint)index] = $"bg:{background}|fg:{textColor}|num:{numberFormat}";
        }

        return result;
    }

    private static T? GetAt<T>(IReadOnlyList<T> items, UInt32Value? index)
        where T : class
    {
        if (index?.Value is not { } value || value > int.MaxValue || value >= items.Count)
        {
            return null;
        }

        return items[(int)value];
    }

    private static string? ColorKey(ColorType? color)
    {
        if (color is null)
        {
            return null;
        }

        if (color.Rgb?.Value is { Length: > 0 } rgb)
        {
            return $"rgb:{rgb}";
        }

        if (color.Indexed?.Value is { } indexed)
        {
            return $"indexed:{indexed}";
        }

        if (color.Theme?.Value is { } theme)
        {
            return $"theme:{theme}:{color.Tint?.Value ?? 0}";
        }

        return color.Auto?.Value == true ? "auto" : null;
    }

    private static string ReadStyle(Cell cell, IReadOnlyDictionary<uint, string> styleLookup)
    {
        var styleIndex = cell.StyleIndex?.Value;
        return styleIndex is null ? string.Empty : styleLookup.GetValueOrDefault(styleIndex.Value, string.Empty);
    }

    private static IReadOnlyDictionary<string, string> ReadDefinedNames(WorkbookPart workbookPart)
    {
        var workbook = workbookPart.Workbook
            ?? throw new InvalidOperationException("Excelブックの構造を読み取れませんでした。");

        return workbook.DefinedNames?.Elements<DefinedName>()
            .GroupBy(name => name.Name?.Value ?? string.Empty, StringComparer.OrdinalIgnoreCase)
            .Where(group => !string.IsNullOrWhiteSpace(group.Key))
            .ToDictionary(group => group.Key, group => string.Join(" | ", group.Select(name => name.Text ?? string.Empty)), StringComparer.OrdinalIgnoreCase)
            ?? new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
    }

    private static MacroSnapshot ReadMacro(WorkbookPart workbookPart, bool includeSource)
    {
        if (workbookPart.VbaProjectPart is null)
        {
            return new(false, false, false, []);
        }

        using var stream = workbookPart.VbaProjectPart.GetStream(FileMode.Open, FileAccess.Read);
        using var memory = new MemoryStream();
        stream.CopyTo(memory);
        var bytes = memory.ToArray();
        var searchable = Encoding.Latin1.GetString(bytes);
        var protectedProject = searchable.Contains("DPB=", StringComparison.OrdinalIgnoreCase) ||
            searchable.Contains("GC=", StringComparison.OrdinalIgnoreCase);

        if (!includeSource)
        {
            return new(true, false, protectedProject, []);
        }

        try
        {
            var source = VbaProjectSourceReader.Read(bytes);
            return new(
                true,
                source.Modules.Count > 0,
                protectedProject,
                source.Modules
                    .Select(module => new MacroModuleSnapshot(
                        module.Name,
                        module.StreamName,
                        module.SourceText,
                        CountSourceLines(module.SourceText)))
                    .ToList());
        }
        catch
        {
            return new(true, false, protectedProject, []);
        }
    }

    private static void CompareStructure(WorkbookSnapshot source, WorkbookSnapshot target, ICollection<DiffRecord> records)
    {
        var sourceNames = source.SheetsByName.Keys.ToHashSet(StringComparer.OrdinalIgnoreCase);
        var targetNames = target.SheetsByName.Keys.ToHashSet(StringComparer.OrdinalIgnoreCase);
        var removed = sourceNames.Except(targetNames, StringComparer.OrdinalIgnoreCase).ToList();
        var added = targetNames.Except(sourceNames, StringComparer.OrdinalIgnoreCase).ToList();
        var renamePairs = MatchRenamedSheets(source, target, removed, added);

        foreach (var (oldName, newName) in renamePairs)
        {
            records.Add(new(DiffCategory.Structure, DiffKind.Renamed, newName, string.Empty, "シート名", oldName, newName, "シート構成が変更されました。"));
            removed.Remove(oldName);
            added.Remove(newName);
        }

        foreach (var name in removed)
        {
            records.Add(new(DiffCategory.Structure, DiffKind.Removed, name, string.Empty, "シート", name, string.Empty, "比較先にシートがありません。"));
        }

        foreach (var name in added)
        {
            records.Add(new(DiffCategory.Structure, DiffKind.Added, name, string.Empty, "シート", string.Empty, name, "比較先でシートが追加されました。"));
        }

        foreach (var name in sourceNames.Intersect(targetNames, StringComparer.OrdinalIgnoreCase))
        {
            var left = source.SheetsByName[name];
            var right = target.SheetsByName[name];
            AddStructureChange(records, name, "シート順", left.Position.ToString(CultureInfo.InvariantCulture), right.Position.ToString(CultureInfo.InvariantCulture));
            AddStructureChange(records, name, "シート非表示状態", left.State, right.State);
            AddStructureChange(records, name, "シート保護", ShortHash(left.ProtectionSignature), ShortHash(right.ProtectionSignature));
            CompareStringSets(records, name, "非表示行", left.HiddenRows, right.HiddenRows);
            CompareStringSets(records, name, "非表示列", left.HiddenColumns, right.HiddenColumns);
            CompareStringSets(records, name, "入力規則", left.DataValidations, right.DataValidations);
            CompareColumnStructure(records, name, left, right);
        }
    }

    private static void CompareColumnStructure(
        ICollection<DiffRecord> records,
        string sheetName,
        SheetSnapshot source,
        SheetSnapshot target)
    {
        var matchedTargets = new HashSet<uint>();
        foreach (var sourceColumn in source.ColumnSignatures.Keys.Order())
        {
            var sourceSignature = source.ColumnSignatures[sourceColumn];
            if (string.IsNullOrEmpty(sourceSignature))
            {
                continue;
            }

            var sameColumn = target.ColumnSignatures.TryGetValue(sourceColumn, out var targetSignature) &&
                sourceSignature.Equals(targetSignature, StringComparison.Ordinal);
            if (sameColumn)
            {
                matchedTargets.Add(sourceColumn);
                continue;
            }

            var movedTarget = target.ColumnSignatures
                .Where(pair => !matchedTargets.Contains(pair.Key) && pair.Value.Equals(sourceSignature, StringComparison.Ordinal))
                .Select(pair => (uint?)pair.Key)
                .FirstOrDefault();

            if (movedTarget is { } targetColumn)
            {
                matchedTargets.Add(targetColumn);
                records.Add(new(
                    DiffCategory.Structure,
                    DiffKind.Changed,
                    sheetName,
                    $"{ColumnName(sourceColumn)} -> {ColumnName(targetColumn)}",
                    "列位置",
                    ColumnName(sourceColumn),
                    ColumnName(targetColumn),
                    "同じ内容の列が別位置へ移動しました。セル比較ではこの列ズレを補正しています。"));
            }
            else if (!target.ColumnSignatures.Values.Contains(sourceSignature, StringComparer.Ordinal))
            {
                records.Add(new(
                    DiffCategory.Structure,
                    DiffKind.Removed,
                    sheetName,
                    ColumnName(sourceColumn),
                    "列",
                    ColumnName(sourceColumn),
                    string.Empty,
                    "比較先で列内容が削除された可能性があります。"));
            }
        }

        foreach (var targetColumn in target.ColumnSignatures.Keys.Order())
        {
            if (matchedTargets.Contains(targetColumn))
            {
                continue;
            }

            var signature = target.ColumnSignatures[targetColumn];
            if (!string.IsNullOrEmpty(signature) && !source.ColumnSignatures.Values.Contains(signature, StringComparer.Ordinal))
            {
                records.Add(new(
                    DiffCategory.Structure,
                    DiffKind.Added,
                    sheetName,
                    ColumnName(targetColumn),
                    "列",
                    string.Empty,
                    ColumnName(targetColumn),
                    "比較先で列内容が追加された可能性があります。"));
            }
        }
    }

    private static void CompareSheets(
        WorkbookSnapshot source,
        WorkbookSnapshot target,
        ComparisonOptions options,
        ICollection<DiffRecord> records)
    {
        foreach (var sheetName in source.SheetsByName.Keys.Intersect(target.SheetsByName.Keys, StringComparer.OrdinalIgnoreCase))
        {
            var left = source.SheetsByName[sheetName];
            var right = target.SheetsByName[sheetName];
            CompareSheetCells(left, right, options, records);
        }
    }

    private static void CompareSheetCells(
        SheetSnapshot source,
        SheetSnapshot target,
        ComparisonOptions options,
        ICollection<DiffRecord> records)
    {
        var leftRows = source.CellsByRow.Keys.Order().ToList();
        var rightRows = target.CellsByRow.Keys.Order().ToList();
        var i = 0;
        var j = 0;

        while (i < leftRows.Count || j < rightRows.Count)
        {
            if (i >= leftRows.Count)
            {
                AddRowBlock(records, DiffKind.Added, target.Name, target.CellsByRow, rightRows.Skip(j), target: true);
                break;
            }

            if (j >= rightRows.Count)
            {
                AddRowBlock(records, DiffKind.Removed, source.Name, source.CellsByRow, leftRows.Skip(i), target: false);
                break;
            }

            var leftRow = leftRows[i];
            var rightRow = rightRows[j];
            var leftSignature = source.RowSignatures.GetValueOrDefault(leftRow, string.Empty);
            var rightSignature = target.RowSignatures.GetValueOrDefault(rightRow, string.Empty);

            if (leftSignature == rightSignature)
            {
                CompareAlignedRows(source, target, leftRow, rightRow, options, records);
                i++;
                j++;
                continue;
            }

            var targetMatch = FindMatchingRow(leftSignature, target.RowSignatures, rightRows, j + 1);
            if (targetMatch >= 0)
            {
                AddRowBlock(records, DiffKind.Added, target.Name, target.CellsByRow, rightRows.Skip(j).Take(targetMatch - j), target: true);
                j = targetMatch;
                continue;
            }

            var sourceMatch = FindMatchingRow(rightSignature, source.RowSignatures, leftRows, i + 1);
            if (sourceMatch >= 0)
            {
                AddRowBlock(records, DiffKind.Removed, source.Name, source.CellsByRow, leftRows.Skip(i).Take(sourceMatch - i), target: false);
                i = sourceMatch;
                continue;
            }

            CompareAlignedRows(source, target, leftRow, rightRow, options, records);
            i++;
            j++;
        }
    }

    private static void CompareAlignedRows(
        SheetSnapshot source,
        SheetSnapshot target,
        uint sourceRow,
        uint targetRow,
        ComparisonOptions options,
        ICollection<DiffRecord> records)
    {
        var leftCells = source.CellsByRow[sourceRow].Values.OrderBy(cell => cell.Column).ToList();
        var rightCells = target.CellsByRow[targetRow].Values.OrderBy(cell => cell.Column).ToList();
        var i = 0;
        var j = 0;

        while (i < leftCells.Count || j < rightCells.Count)
        {
            if (i >= leftCells.Count)
            {
                AddCellBlock(records, DiffKind.Added, target.Name, rightCells.Skip(j), target: true, options);
                break;
            }

            if (j >= rightCells.Count)
            {
                AddCellBlock(records, DiffKind.Removed, source.Name, leftCells.Skip(i), target: false, options);
                break;
            }

            var left = leftCells[i];
            var right = rightCells[j];
            if (SameCellIdentity(left, right))
            {
                CompareAlignedCells(target.Name, left, right, options, records);
                i++;
                j++;
                continue;
            }

            var targetMatch = FindMatchingCell(left, rightCells, j + 1);
            if (targetMatch >= 0)
            {
                AddCellBlock(records, DiffKind.Added, target.Name, rightCells.Skip(j).Take(targetMatch - j), target: true, options);
                j = targetMatch;
                continue;
            }

            var sourceMatch = FindMatchingCell(right, leftCells, i + 1);
            if (sourceMatch >= 0)
            {
                AddCellBlock(records, DiffKind.Removed, source.Name, leftCells.Skip(i).Take(sourceMatch - i), target: false, options);
                i = sourceMatch;
                continue;
            }

            CompareAlignedCells(target.Name, left, right, options, records);
            i++;
            j++;
        }
    }

    private static void CompareAlignedCells(
        string sheetName,
        CellSnapshot left,
        CellSnapshot right,
        ComparisonOptions options,
        ICollection<DiffRecord> records)
    {
        var address = right.Address;
        if (options.IncludeData)
        {
            AddCellChange(records, sheetName, address, "値", left.Value, right.Value);
            AddCellChange(records, sheetName, address, "数式", left.Formula, right.Formula);
        }

        if (options.IncludeFormatting)
        {
            AddFormattingChange(records, sheetName, address, left.StyleSignature, right.StyleSignature);
        }
    }

    private static bool SameCellIdentity(CellSnapshot left, CellSnapshot right)
    {
        if (!string.IsNullOrEmpty(left.DataSignature) || !string.IsNullOrEmpty(right.DataSignature))
        {
            return string.Equals(left.DataSignature, right.DataSignature, StringComparison.Ordinal);
        }

        return string.Equals(left.StyleSignature, right.StyleSignature, StringComparison.Ordinal);
    }

    private static int FindMatchingCell(CellSnapshot needle, IReadOnlyList<CellSnapshot> cells, int startIndex)
    {
        var max = Math.Min(cells.Count, startIndex + ShiftLookAhead);
        for (var index = startIndex; index < max; index++)
        {
            if (SameCellIdentity(needle, cells[index]))
            {
                return index;
            }
        }

        return -1;
    }

    private static void AddCellBlock(
        ICollection<DiffRecord> records,
        DiffKind kind,
        string sheet,
        IEnumerable<CellSnapshot> cells,
        bool target,
        ComparisonOptions options)
    {
        foreach (var cell in cells)
        {
            if (options.IncludeData && !string.IsNullOrEmpty(cell.DataSignature))
            {
                records.Add(new(
                    DiffCategory.Cell,
                    kind,
                    sheet,
                    cell.Address,
                    "セル",
                    target ? string.Empty : cell.DataSignature,
                    target ? cell.DataSignature : string.Empty,
                    kind == DiffKind.Added ? "列挿入またはセル追加を検出しました。" : "列削除またはセル削除を検出しました。"));
            }

            if (options.IncludeFormatting && !string.IsNullOrEmpty(cell.StyleSignature))
            {
                records.Add(new(
                    DiffCategory.Formatting,
                    kind,
                    sheet,
                    cell.Address,
                    "書式",
                    target ? string.Empty : cell.StyleSignature,
                    target ? cell.StyleSignature : string.Empty,
                    kind == DiffKind.Added ? "追加セルの書式を検出しました。" : "削除セルの書式を検出しました。"));
            }
        }
    }

    private static void CompareMacros(WorkbookSnapshot source, WorkbookSnapshot target, ICollection<DiffRecord> records)
    {
        if (!source.Metadata.Extension.Equals(".xlsm", StringComparison.OrdinalIgnoreCase) ||
            !target.Metadata.Extension.Equals(".xlsm", StringComparison.OrdinalIgnoreCase))
        {
            records.Add(new(
                DiffCategory.Macro,
                DiffKind.Changed,
                string.Empty,
                string.Empty,
                "マクロ比較",
                source.Metadata.Extension,
                target.Metadata.Extension,
                "マクロ比較は両方のファイルが .xlsm の場合のみ詳細実行します。"));
            return;
        }

        if (source.Macro.HasMacro != target.Macro.HasMacro)
        {
            records.Add(new(DiffCategory.Macro, DiffKind.Changed, string.Empty, string.Empty, "マクロ有無", source.Macro.HasMacro ? "あり" : "なし", target.Macro.HasMacro ? "あり" : "なし", "VBAプロジェクトの有無が変わりました。"));
        }

        if (!source.Macro.HasMacro && !target.Macro.HasMacro)
        {
            return;
        }

        if (source.Macro.IsPasswordProtected != target.Macro.IsPasswordProtected)
        {
            records.Add(new(DiffCategory.Macro, DiffKind.Changed, string.Empty, string.Empty, "VBA保護", source.Macro.IsPasswordProtected ? "あり" : "なし", target.Macro.IsPasswordProtected ? "あり" : "なし", "VBAプロジェクト保護の痕跡が変わりました。"));
        }

        var leftModules = source.Macro.Modules.ToDictionary(module => module.Name, StringComparer.OrdinalIgnoreCase);
        var rightModules = target.Macro.Modules.ToDictionary(module => module.Name, StringComparer.OrdinalIgnoreCase);

        foreach (var name in leftModules.Keys.Except(rightModules.Keys, StringComparer.OrdinalIgnoreCase))
        {
            var module = leftModules[name];
            records.Add(new(
                DiffCategory.Macro,
                DiffKind.Removed,
                string.Empty,
                string.Empty,
                name,
                $"{module.LineCount}行",
                string.Empty,
                "VBAモジュールが削除されました。",
                PreviewCode(module.SourceText),
                string.Empty));
        }

        foreach (var name in rightModules.Keys.Except(leftModules.Keys, StringComparer.OrdinalIgnoreCase))
        {
            var module = rightModules[name];
            records.Add(new(
                DiffCategory.Macro,
                DiffKind.Added,
                string.Empty,
                string.Empty,
                name,
                string.Empty,
                $"{module.LineCount}行",
                "VBAモジュールが追加されました。",
                string.Empty,
                PreviewCode(module.SourceText)));
        }

        foreach (var name in leftModules.Keys.Intersect(rightModules.Keys, StringComparer.OrdinalIgnoreCase))
        {
            var left = leftModules[name];
            var right = rightModules[name];
            if (!left.SourceText.Equals(right.SourceText, StringComparison.Ordinal))
            {
                var changedLines = CountChangedLines(left.SourceText, right.SourceText);
                records.Add(new(DiffCategory.Macro, DiffKind.Changed, string.Empty, string.Empty, name, $"{left.LineCount}行", $"{right.LineCount}行", $"コードが変更されました。変更行数目安: {changedLines:N0}行"));
                AddMacroLineDiffs(records, name, left.SourceText, right.SourceText);
            }
        }
    }

    private static void CompareDefinedNames(WorkbookSnapshot source, WorkbookSnapshot target, ICollection<DiffRecord> records)
    {
        var names = source.DefinedNames.Keys.Union(target.DefinedNames.Keys, StringComparer.OrdinalIgnoreCase).Order(StringComparer.OrdinalIgnoreCase);
        foreach (var name in names)
        {
            var hasLeft = source.DefinedNames.TryGetValue(name, out var left);
            var hasRight = target.DefinedNames.TryGetValue(name, out var right);
            var kind = (hasLeft, hasRight) switch
            {
                (false, true) => DiffKind.Added,
                (true, false) => DiffKind.Removed,
                _ => DiffKind.Changed
            };

            if (!string.Equals(left, right, StringComparison.Ordinal))
            {
                records.Add(new(DiffCategory.Metadata, kind, string.Empty, string.Empty, $"名前定義:{name}", left ?? string.Empty, right ?? string.Empty, "名前定義の参照範囲または式が変更されました。"));
            }
        }
    }

    private static void AddStructureChange(ICollection<DiffRecord> records, string sheet, string item, string source, string target)
    {
        if (!string.Equals(source, target, StringComparison.Ordinal))
        {
            records.Add(new(DiffCategory.Structure, DiffKind.Changed, sheet, string.Empty, item, source, target, $"{item}が変更されました。"));
        }
    }

    private static void CompareStringSets(
        ICollection<DiffRecord> records,
        string sheet,
        string item,
        IReadOnlyList<string> source,
        IReadOnlyList<string> target)
    {
        foreach (var value in source.Except(target, StringComparer.OrdinalIgnoreCase).Take(MaxPerShiftBlock))
        {
            records.Add(new(DiffCategory.Structure, DiffKind.Removed, sheet, string.Empty, item, value, string.Empty, $"{item}が削除されました。"));
        }

        foreach (var value in target.Except(source, StringComparer.OrdinalIgnoreCase).Take(MaxPerShiftBlock))
        {
            records.Add(new(DiffCategory.Structure, DiffKind.Added, sheet, string.Empty, item, string.Empty, value, $"{item}が追加されました。"));
        }
    }

    private static void AddCellChange(ICollection<DiffRecord> records, string sheet, string address, string item, string source, string target)
    {
        if (!string.Equals(source, target, StringComparison.Ordinal))
        {
            records.Add(new(DiffCategory.Cell, DiffKind.Changed, sheet, address, item, source, target, $"{item}が変更されました。"));
        }
    }

    private static void AddFormattingChange(ICollection<DiffRecord> records, string sheet, string address, string source, string target)
    {
        if (!string.Equals(source, target, StringComparison.Ordinal))
        {
            records.Add(new(DiffCategory.Formatting, DiffKind.Changed, sheet, address, "書式", source, target, "背景色、文字色、または表示形式が変更されました。"));
        }
    }

    private static int FindMatchingRow(
        string signature,
        IReadOnlyDictionary<uint, string> signatures,
        IReadOnlyList<uint> rows,
        int startIndex)
    {
        if (string.IsNullOrEmpty(signature))
        {
            return -1;
        }

        var max = Math.Min(rows.Count, startIndex + ShiftLookAhead);
        for (var index = startIndex; index < max; index++)
        {
            if (signatures.GetValueOrDefault(rows[index], string.Empty) == signature)
            {
                return index;
            }
        }

        return -1;
    }

    private static void AddRowBlock(
        ICollection<DiffRecord> records,
        DiffKind kind,
        string sheet,
        IReadOnlyDictionary<uint, IReadOnlyDictionary<uint, CellSnapshot>> rows,
        IEnumerable<uint> rowNumbers,
        bool target)
    {
        var emitted = 0;
        foreach (var rowNumber in rowNumbers)
        {
            foreach (var cell in rows[rowNumber].Values.OrderBy(cell => cell.Column))
            {
                if (emitted++ >= MaxPerShiftBlock)
                {
                    records.Add(new(DiffCategory.Cell, kind, sheet, string.Empty, "行ブロック", string.Empty, string.Empty, "大量の行差分があるため、一部のみ表示しています。"));
                    return;
                }

                records.Add(new(
                    DiffCategory.Cell,
                    kind,
                    sheet,
                    cell.Address,
                    "セル",
                    target ? string.Empty : cell.DataSignature,
                    target ? cell.DataSignature : string.Empty,
                    kind == DiffKind.Added ? "行挿入またはセル追加を検出しました。" : "行削除またはセル削除を検出しました。"));
            }
        }
    }

    private static IReadOnlyList<(string Source, string Target)> MatchRenamedSheets(
        WorkbookSnapshot source,
        WorkbookSnapshot target,
        IReadOnlyList<string> removed,
        IReadOnlyList<string> added)
    {
        var pairs = new List<(string Source, string Target)>();
        var usedTargets = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        foreach (var oldName in removed)
        {
            var left = source.SheetsByName[oldName];
            var bestScore = 0.0;
            string? bestName = null;
            foreach (var newName in added.Where(name => !usedTargets.Contains(name)))
            {
                var score = SheetSimilarity(left, target.SheetsByName[newName]);
                if (score > bestScore)
                {
                    bestScore = score;
                    bestName = newName;
                }
            }

            if (bestName is not null && bestScore >= 0.7)
            {
                pairs.Add((oldName, bestName));
                usedTargets.Add(bestName);
            }
        }

        return pairs;
    }

    private static double SheetSimilarity(SheetSnapshot left, SheetSnapshot right)
    {
        var leftSet = left.RowSignatures.Values.Where(value => value.Length > 0).Take(200).ToHashSet(StringComparer.Ordinal);
        var rightSet = right.RowSignatures.Values.Where(value => value.Length > 0).Take(200).ToHashSet(StringComparer.Ordinal);
        if (leftSet.Count == 0 && rightSet.Count == 0)
        {
            return 1;
        }

        var intersection = leftSet.Intersect(rightSet, StringComparer.Ordinal).Count();
        var union = leftSet.Union(rightSet, StringComparer.Ordinal).Count();
        return union == 0 ? 0 : (double)intersection / union;
    }

    private static string BuildRowSignature(IReadOnlyDictionary<uint, CellSnapshot> row)
    {
        if (row.Count == 0)
        {
            return string.Empty;
        }

        var builder = new StringBuilder();
        foreach (var cell in row.OrderBy(pair => pair.Key).Select(pair => pair.Value))
        {
            if (string.IsNullOrEmpty(cell.DataSignature))
            {
                continue;
            }

            builder.Append(cell.DataSignature);
            builder.Append('\u001f');
        }

        return builder.ToString();
    }

    private static string BuildColumnSignature(IEnumerable<CellSnapshot> column)
    {
        var builder = new StringBuilder();
        foreach (var cell in column.OrderBy(cell => cell.Row))
        {
            if (string.IsNullOrEmpty(cell.DataSignature))
            {
                continue;
            }

            builder.Append(cell.DataSignature);
            builder.Append('\u001f');
        }

        return builder.ToString();
    }

    private static IReadOnlyList<string> ReadHiddenRows(Worksheet worksheet)
    {
        return worksheet.Descendants<Row>()
            .Where(row => row.Hidden?.Value == true && row.RowIndex?.Value is not null)
            .Select(row => row.RowIndex!.Value.ToString(CultureInfo.InvariantCulture))
            .ToList();
    }

    private static IReadOnlyList<string> ReadHiddenColumns(Worksheet worksheet)
    {
        return worksheet.Descendants<Column>()
            .Where(column => column.Hidden?.Value == true)
            .Select(column => $"{column.Min?.Value}-{column.Max?.Value}")
            .ToList();
    }

    private static IReadOnlyList<string> ReadDataValidations(Worksheet worksheet)
    {
        return worksheet.Descendants<DataValidation>()
            .Select(validation =>
                $"{validation.SequenceOfReferences?.InnerText}|{validation.Type?.Value}|{validation.Operator?.Value}|{validation.Formula1?.Text}|{validation.Formula2?.Text}")
            .Order(StringComparer.Ordinal)
            .ToList();
    }

    private static (uint Column, uint Row) ParseCellReference(ReadOnlySpan<char> reference)
    {
        uint column = 0;
        uint row = 0;

        foreach (var ch in reference)
        {
            if (ch is >= 'A' and <= 'Z')
            {
                column = column * 26 + (uint)(ch - 'A' + 1);
                continue;
            }

            if (ch is >= 'a' and <= 'z')
            {
                column = column * 26 + (uint)(ch - 'a' + 1);
                continue;
            }

            if (ch is >= '0' and <= '9')
            {
                row = row * 10 + (uint)(ch - '0');
            }
        }

        return (column, row);
    }

    private static string ColumnName(uint column)
    {
        Span<char> buffer = stackalloc char[8];
        var index = buffer.Length;
        while (column > 0)
        {
            column--;
            buffer[--index] = (char)('A' + column % 26);
            column /= 26;
        }

        return new string(buffer[index..]);
    }

    private static int CountSourceLines(string text) =>
        string.IsNullOrEmpty(text) ? 0 : text.Count(character => character == '\n') + 1;

    private static int CountChangedLines(string left, string right)
    {
        var leftLines = SplitLines(left);
        var rightLines = SplitLines(right);
        var common = CountCommonLines(leftLines, rightLines);
        return Math.Max(0, leftLines.Length - common) + Math.Max(0, rightLines.Length - common);
    }

    private static void AddMacroLineDiffs(
        ICollection<DiffRecord> records,
        string moduleName,
        string sourceText,
        string targetText)
    {
        var sourceLines = SplitLines(sourceText);
        var targetLines = SplitLines(targetText);
        var i = 0;
        var j = 0;
        var emitted = 0;

        while ((i < sourceLines.Length || j < targetLines.Length) && emitted < MaxMacroLineDiffsPerModule)
        {
            if (i < sourceLines.Length &&
                j < targetLines.Length &&
                sourceLines[i].Equals(targetLines[j], StringComparison.Ordinal))
            {
                i++;
                j++;
                continue;
            }

            if (i + 1 < sourceLines.Length &&
                j < targetLines.Length &&
                sourceLines[i + 1].Equals(targetLines[j], StringComparison.Ordinal))
            {
                records.Add(MacroLineRecord(moduleName, DiffKind.Removed, i + 1, null, sourceLines[i], string.Empty));
                i++;
                emitted++;
                continue;
            }

            if (j + 1 < targetLines.Length &&
                i < sourceLines.Length &&
                sourceLines[i].Equals(targetLines[j + 1], StringComparison.Ordinal))
            {
                records.Add(MacroLineRecord(moduleName, DiffKind.Added, null, j + 1, string.Empty, targetLines[j]));
                j++;
                emitted++;
                continue;
            }

            if (i < sourceLines.Length && j < targetLines.Length)
            {
                records.Add(MacroLineRecord(moduleName, DiffKind.Changed, i + 1, j + 1, sourceLines[i], targetLines[j]));
                i++;
                j++;
                emitted++;
                continue;
            }

            if (i < sourceLines.Length)
            {
                records.Add(MacroLineRecord(moduleName, DiffKind.Removed, i + 1, null, sourceLines[i], string.Empty));
                i++;
                emitted++;
                continue;
            }

            records.Add(MacroLineRecord(moduleName, DiffKind.Added, null, j + 1, string.Empty, targetLines[j]));
            j++;
            emitted++;
        }

        if (i < sourceLines.Length || j < targetLines.Length)
        {
            records.Add(new(
                DiffCategory.Macro,
                DiffKind.Changed,
                string.Empty,
                string.Empty,
                $"{moduleName} 行差分",
                string.Empty,
                string.Empty,
                $"行差分が多いため、先頭 {MaxMacroLineDiffsPerModule:N0} 件のみ表示しています。"));
        }
    }

    private static DiffRecord MacroLineRecord(
        string moduleName,
        DiffKind kind,
        int? sourceLine,
        int? targetLine,
        string source,
        string target)
    {
        var address = (sourceLine, targetLine) switch
        {
            (not null, not null) => $"L{sourceLine} -> L{targetLine}",
            (not null, null) => $"L{sourceLine}",
            (null, not null) => $"L{targetLine}",
            _ => string.Empty
        };

        return new(
            DiffCategory.Macro,
            kind,
            string.Empty,
            address,
            $"{moduleName} 行差分",
            Shorten(source),
            Shorten(target),
            "VBAコード行の差分です。",
            source,
            target);
    }

    private static string PreviewCode(string text)
    {
        var lines = SplitLines(text)
            .Where(line => !string.IsNullOrWhiteSpace(line))
            .Take(12);

        return string.Join('\n', lines);
    }

    private static string Shorten(string value)
    {
        const int maxLength = 160;
        if (value.Length <= maxLength)
        {
            return value;
        }

        return value[..maxLength] + "...";
    }

    private static string[] SplitLines(string text) =>
        text.Replace("\r\n", "\n", StringComparison.Ordinal).Replace('\r', '\n').Split('\n');

    private static int CountCommonLines(string[] left, string[] right)
    {
        if ((long)left.Length * right.Length > 400_000)
        {
            return left.Zip(right).Count(pair => pair.First.Equals(pair.Second, StringComparison.Ordinal));
        }

        var previous = new int[right.Length + 1];
        var current = new int[right.Length + 1];
        for (var i = 1; i <= left.Length; i++)
        {
            for (var j = 1; j <= right.Length; j++)
            {
                current[j] = left[i - 1].Equals(right[j - 1], StringComparison.Ordinal)
                    ? previous[j - 1] + 1
                    : Math.Max(previous[j], current[j - 1]);
            }

            (previous, current) = (current, previous);
            Array.Clear(current);
        }

        return previous[right.Length];
    }

    private static string ShortHash(string value)
    {
        if (string.IsNullOrEmpty(value))
        {
            return string.Empty;
        }

        var hash = SHA256.HashData(Encoding.UTF8.GetBytes(value));
        return Convert.ToHexString(hash.AsSpan(0, 4));
    }
}
