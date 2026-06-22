using System.Globalization;
using System.Text;
using CsvHelper;
using CsvHelper.Configuration;
using ExcelCompare.Models;

namespace ExcelCompare.Services;

public sealed class CsvExportService
{
    public byte[] BuildDiffCsv(IEnumerable<DiffRecord> records)
    {
        using var memory = new MemoryStream();
        memory.Write(Encoding.UTF8.GetPreamble());
        using (var writer = new StreamWriter(memory, new UTF8Encoding(encoderShouldEmitUTF8Identifier: false), leaveOpen: true))
        using (var csv = new CsvWriter(writer, new CsvConfiguration(CultureInfo.InvariantCulture)
        {
            HasHeaderRecord = true
        }))
        {
            csv.WriteField("カテゴリ");
            csv.WriteField("種別");
            csv.WriteField("シート");
            csv.WriteField("アドレス");
            csv.WriteField("項目");
            csv.WriteField("比較元");
            csv.WriteField("比較先");
            csv.WriteField("詳細");
            csv.NextRecord();

            foreach (var record in records)
            {
                csv.WriteField(ToLabel(record.Category));
                csv.WriteField(ToLabel(record.Kind));
                csv.WriteField(record.SheetName);
                csv.WriteField(record.Address);
                csv.WriteField(record.Item);
                csv.WriteField(record.SourceValue);
                csv.WriteField(record.TargetValue);
                csv.WriteField(record.Detail);
                csv.NextRecord();
            }
        }

        return memory.ToArray();
    }

    private static string ToLabel(DiffCategory category) => category switch
    {
        DiffCategory.Structure => "シート構成",
        DiffCategory.Cell => "セル詳細",
        DiffCategory.Formatting => "書式",
        DiffCategory.Macro => "マクロ",
        DiffCategory.Metadata => "その他メタデータ",
        _ => category.ToString()
    };

    private static string ToLabel(DiffKind kind) => kind switch
    {
        DiffKind.Added => "追加",
        DiffKind.Removed => "削除",
        DiffKind.Changed => "変更",
        DiffKind.Renamed => "名称変更",
        _ => kind.ToString()
    };
}
