namespace ExcelCompare.Models;

public enum DiffCategory
{
    Structure,
    Cell,
    Formatting,
    Macro,
    Metadata
}

public enum DiffKind
{
    Added,
    Removed,
    Changed,
    Renamed
}

public sealed record ComparisonOptions(
    bool IncludeData,
    bool IncludeFormatting,
    bool IncludeStructure,
    bool IncludeMacros);

public sealed record CompareProgress(string Status, int Percent);

public sealed record FileMetadata(
    string FileName,
    long SizeBytes,
    string Extension,
    int SheetCount,
    IReadOnlyList<string> SheetNames,
    bool HasMacro);

public sealed record DiffRecord(
    DiffCategory Category,
    DiffKind Kind,
    string SheetName,
    string Address,
    string Item,
    string SourceValue,
    string TargetValue,
    string Detail,
    string SourcePreview = "",
    string TargetPreview = "");

public sealed record DiffSummary(
    int StructureCount,
    int CellCount,
    int FormattingCount,
    int MacroCount,
    int MetadataCount)
{
    public int Total => StructureCount + CellCount + FormattingCount + MacroCount + MetadataCount;
}

public sealed record CompareResult(
    FileMetadata Source,
    FileMetadata Target,
    ComparisonOptions Options,
    DiffSummary Summary,
    IReadOnlyList<DiffRecord> Records,
    IReadOnlyList<DiffRecord> DisplayRecords);

internal sealed record WorkbookSnapshot(
    FileMetadata Metadata,
    IReadOnlyList<SheetSnapshot> Sheets,
    IReadOnlyDictionary<string, string> DefinedNames,
    MacroSnapshot Macro)
{
    public IReadOnlyDictionary<string, SheetSnapshot> SheetsByName { get; } =
        Sheets.ToDictionary(sheet => sheet.Name, StringComparer.OrdinalIgnoreCase);
}

internal sealed record SheetSnapshot(
    string Name,
    uint SheetId,
    int Position,
    string State,
    IReadOnlyDictionary<string, CellSnapshot> CellsByAddress,
    IReadOnlyDictionary<uint, string> RowSignatures,
    IReadOnlyDictionary<uint, string> ColumnSignatures,
    IReadOnlyDictionary<uint, IReadOnlyDictionary<uint, CellSnapshot>> CellsByRow,
    IReadOnlyList<string> HiddenRows,
    IReadOnlyList<string> HiddenColumns,
    IReadOnlyList<string> DataValidations,
    string ProtectionSignature,
    string ProtectionSummary);

internal sealed record CellSnapshot(
    string Address,
    uint Row,
    uint Column,
    string Value,
    string Formula,
    string StyleSignature)
{
    public string DataSignature => string.IsNullOrEmpty(Formula)
        ? Value
        : $"={Formula}";
}

internal sealed record MacroSnapshot(
    bool HasMacro,
    bool SourceCodeParsed,
    bool IsPasswordProtected,
    IReadOnlyList<MacroModuleSnapshot> Modules);

internal sealed record MacroModuleSnapshot(
    string Name,
    string StreamName,
    string SourceText,
    int LineCount);
