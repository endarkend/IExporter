using DocumentFormat.OpenXml.Spreadsheet;
using ExportWrapper.Common.Models;

namespace ExportWrapper.Common.ColumnConfigFilters;
public class YesNoFilter : IColumnConfigFilter
{
    public string[] ColumnNames { get; set; }
    public (dynamic, CellValues, uint?, string) ApplyFilter(dynamic item, RowConfiguration rowConfig)
    {
        return ((bool)item ? "Yes" : "No", CellValues.String, null, (bool)item ? "Yes" : "No");
    }
}
