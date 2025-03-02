using DocumentFormat.OpenXml.Spreadsheet;
using ExportWrapper.Common.Models;
using ExportWrapper.Enums;

namespace ExportWrapper.Common.ColumnConfigFilters;
public class IntDoubleNumberFilter : IColumnConfigFilter
{
    public string[] ColumnNames { get => throw new System.NotImplementedException(); set => throw new System.NotImplementedException(); }

    public (dynamic, CellValues, uint?, string) ApplyFilter(dynamic item, RowConfiguration rowConfig = null)
    {
        if (item == null)
            return (null, CellValues.Number, (uint)ExcelCellTypes.Default, string.Empty);
        return ((double)item, CellValues.Number, (uint)ExcelCellTypes.Default, ((double)item).ToString());
    }
}
