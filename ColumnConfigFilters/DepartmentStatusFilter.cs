
using DocumentFormat.OpenXml.Spreadsheet;
using ExportWrapper.Common.Models;

namespace ExportWrapper.Common.ColumnConfigFilters;
public class DepartmentStatusFilter : IColumnConfigFilter
{
    public string[] ColumnNames { get; set; }

    public (dynamic, CellValues, uint?, string) ApplyFilter(dynamic item, RowConfiguration rowConfig)
    {
        return ((bool)item ? "locked" : "unlocked", CellValues.String, null, (bool)item ? "locked" : "unlocked");
    }
}
