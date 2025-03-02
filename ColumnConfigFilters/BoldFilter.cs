using DocumentFormat.OpenXml.Spreadsheet;
using ExportWrapper.Enums;

namespace ExportWrapper.Common.ColumnConfigFilters;
public class BoldFilter : IColumnConfigFilter
{
    public string[] ColumnNames { get; set; }

    public (dynamic, CellValues, uint?, string) ApplyFilter(dynamic item, Models.RowConfiguration rowConfig)
    {
        if (rowConfig == null)
            return (item, CellValues.String, (uint)ExcelCellTypes.Bold, item);

        return (item, CellValues.String, rowConfig.BoldStyledCell, item);
    }
}
