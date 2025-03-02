using DocumentFormat.OpenXml;
using System;
using ExportWrapper.Common.ColumnConfigFilters;

namespace ExportWrapper.Common.Models;
public class ColumnConfig
{
    public string ColumnName { get; set; }
    public string? ClassNameMatch { get; set; }
    public DoubleValue Width { get; set; }
    public int PdfColumnTruncateWidth { get; set; }

    public IColumnConfigFilter ColumnConfigFilter { get; set; }

    public ColumnConfig(string columnName, string classNameMatch, DoubleValue width)
    {
        if (string.IsNullOrEmpty(classNameMatch))
            throw new ArgumentException("Export requires a column name");

        ColumnName = columnName;
        ClassNameMatch = classNameMatch;
        Width = width;
        PdfColumnTruncateWidth = (int)Math.Round(width / 1.5);
    }

    public ColumnConfig(string columnName, string classNameMatch, DoubleValue width, IColumnConfigFilter columnConfigFilter)
    {
        if (string.IsNullOrEmpty(classNameMatch))
            throw new ArgumentException("Export requires a column name");

        ColumnName = columnName;
        ClassNameMatch = classNameMatch;
        Width = width;
        ColumnConfigFilter = columnConfigFilter;
        PdfColumnTruncateWidth = (int)Math.Round(width / 1.5);
    }

    public ColumnConfig(string columnName, DoubleValue width, IColumnConfigFilter columnConfigFilter)
    {
        ColumnName = columnName;
        ClassNameMatch = null;
        Width = width;
        ColumnConfigFilter = columnConfigFilter;
        PdfColumnTruncateWidth = (int)Math.Round(width / 1.5);
    }
}
