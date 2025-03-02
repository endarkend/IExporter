using DocumentFormat.OpenXml.Spreadsheet;

namespace ExportWrapper.Common.ColumnConfigFilters;
public interface IColumnConfigFilter
{
    public string[] ColumnNames { get; set; }
    (dynamic, CellValues, uint?, string) ApplyFilter(dynamic item, Models.RowConfiguration rowConfig = null);
}
