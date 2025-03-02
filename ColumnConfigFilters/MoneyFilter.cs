using DocumentFormat.OpenXml.Spreadsheet;
using ExportWrapper.Common.Models;
using ExportWrapper.Enums;

namespace ExportWrapper.Common.ColumnConfigFilters;
public class MoneyFilter : IColumnConfigFilter
{
    public string[] ColumnNames { get; set; }
    public (dynamic, CellValues, uint?, string) ApplyFilter(dynamic item, RowConfiguration rowConfig)
    {
        if (item == null)
            return (null, CellValues.Number, (uint)ExcelCellTypes.Currency, string.Empty);

        if (rowConfig == null)
            return ((double)item, CellValues.Number, (uint)ExcelCellTypes.Currency, ((double)item).ToString("C"));

        return ((double)item, CellValues.Number, rowConfig.MoneyStyledCell, ((double)item).ToString("C"));
    }
}
