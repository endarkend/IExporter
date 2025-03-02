using FastMember;
using PdfSharp.Charting;
using PdfSharp.Drawing;
using PdfSharp.Pdf;
using ExportWrapper.Common.Models;
using ZFreedom.Common.Models;

namespace ExportWrapper.Common.Builders;

/// <summary>
/// Code I have refactored a bit,
/// </summary>
/// <typeparam name="T"></typeparam>
public class PDFDocumentBuilder<T> where T : class
{
    IReadOnlyCollection<ColumnConfig> _columnConfigs;
    ReportStructure _reportStructure;
    TypeAccessor _typeAccessor = TypeAccessor.Create(typeof(T));
    Dictionary<string, string> _properties = typeof(T).GetProperties().ToDictionary(k => k.Name.ToLower(), v => v.Name);
    XGraphics _gfx;
    double _pageWidth = ReportDefaults.A4PortraitWidth;
    double _pageLength = ReportDefaults.A4PortraitHeight;
    int pageNumber = 1;



    public byte[] BuildDocument(Dictionary<string, List<T>> data, IReadOnlyCollection<ColumnConfig> columnConfigs, string reportName)
    {
        double colWidth = columnConfigs.Sum(c => c.Width);

        if (colWidth > 160)
        {
            _pageWidth = ReportDefaults.A4LandscapeWidth;
            _pageLength = ReportDefaults.A4LandscapeHeight;
        }

        _reportStructure = new ReportStructure
        {
            Sections = new List<ReportSection>
                {
                    //Charts
                    new ReportSection
                    {
                        FirstPageTextFields = new List<TextField>
                        {
                            new TextField(reportName, new XRect(0,0,_pageWidth, 12))
                        },
                        Columns = columnConfigs.Select(i => new ReportColumns(i.ColumnName, i.ClassNameMatch, i.Width)).ToList()
                    }
                }
        };

        _columnConfigs = columnConfigs;

        return GenerateReport(data, reportName);
    }   


    public byte[] BuildDocument(IReadOnlyCollection<T> data, IReadOnlyCollection<ColumnConfig> columnConfigs, string reportName)
    {

        double colWidth =columnConfigs.Sum(c => c.Width);

        if(colWidth > 160) {
            _pageWidth = ReportDefaults.A4LandscapeWidth;
            _pageLength = ReportDefaults.A4LandscapeHeight;
        }

        _reportStructure = new ReportStructure
        {
            Sections = new List<ReportSection>
                {
                    //Charts
                    new ReportSection
                    {
                        FirstPageTextFields = new List<TextField>
                        {
                            new TextField(reportName, new XRect(0,0,_pageWidth, 12))
                        },
                        Columns = columnConfigs.Select(i => new ReportColumns(i.ColumnName, i.ClassNameMatch, i.Width)).ToList()
                    }
                }
        };

        _columnConfigs = columnConfigs;

        return GenerateReport(data.ToList(), reportName);
    }

    private byte[] GenerateReport(List<T> rows, string? reportName = null)
    {
        PdfDocument document = new PdfDocument();
        document.Info.Title = string.IsNullOrEmpty(reportName) ? "Test Report" : reportName;

        ProcessSections(rows, document);
        //Return the pdf as a byte array for processing further on
        using (var stream = new MemoryStream())
        {
            document.Save(stream);
            return stream.ToArray();
        }
    }


    private byte[] GenerateReport(Dictionary<string, List<T>> reportData, string? reportName = null)
    { 
        PdfDocument document = new PdfDocument();
        document.Info.Title = string.IsNullOrEmpty(reportName) ? "Test Report" : reportName;
        foreach(var data in reportData)
        {
            ProcessSections(data.Value.ToList(), document, data.Key);
        }       
        using (var stream = new MemoryStream())
        {
            document.Save(stream);
            return stream.ToArray();
        }
    }




    private void ProcessSections(List<T> rows, PdfDocument document, string pageName = null) {

        foreach (var section in _reportStructure.Sections)
        {
            var totalRows = rows.Count;

            var hasMorePages = true;
            var rowIndex = 0;

            //Pre-calculate all font heights
            section.CalculateFontHeights();

            var columnGrouping = SetupGrouping(rows, section);
            var displayRows = columnGrouping.Item1;
            var groupByValues = columnGrouping.Item2;

            var firstPage = true;
            while (hasMorePages)
            {
                var newPage = SetupPage(firstPage, document, section, pageName);
                
                var rowHeight = newPage.rowHeight;
                var pageWidth = newPage.width;
                var pageHeight = newPage.Height;

                SetupColumnsAndMargins(section, rowHeight, pageWidth);
                hasMorePages = ManageDataAndRows(section, rowHeight, rowIndex, pageHeight, totalRows, displayRows, groupByValues, pageWidth, out var currentIndex);
                rowIndex = currentIndex;

                firstPage = false;
                var positionBotomRight = new XPoint(pageWidth, pageHeight);
                _gfx.DrawString($"page {pageNumber}", section.FirstPageTextFields.First().Font, section.FirstPageTextFields.First().Brush, positionBotomRight, section.FirstPageTextFields.First().StringFormat);
                pageNumber++;
            }
        }
    }

    private ChartFrame GenerateChart(ChartField details)
    {
        var chart = new Chart(details.ChartType);
        foreach (var series in details.Series)
        {
            var s = chart.SeriesCollection.AddSeries();
            s.Name = series.Key;
            s.Add(series.Value);
        }

        chart.XAxis.MajorTickMark = TickMarkType.Outside;
        chart.XAxis.Title.Caption = details.XAxisTitle;

        chart.YAxis.MajorTickMark = TickMarkType.Outside;
        chart.YAxis.Title.Caption = details.YAxisTitle;
        chart.YAxis.HasMajorGridlines = true;

        chart.PlotArea.LineFormat.Color = XColors.DarkGray;
        chart.PlotArea.LineFormat.Width = 1;
        chart.PlotArea.LineFormat.Visible = true;

        chart.Legend.Docking = DockingType.Bottom;
        chart.Legend.LineFormat.Visible = true;

        foreach (var xseries in details.XSeries)
        {
            var s = chart.XValues.AddXSeries();
            s.Add(xseries);
        }

        var chartFrame = new ChartFrame();
        chartFrame.Location = new XPoint(details.Position.X, details.Position.Y);
        chartFrame.Size = new XSize(details.Position.Width, details.Position.Height);
        chartFrame.Add(chart);

        return chartFrame;
    }

    private (double rowHeight, double width, double Height) SetupPage(bool firstPage, PdfDocument document, ReportSection section, string pageHeading = null)
    {
        // Create an empty page
        PdfPage page = document.AddPage();
        page.Size = PdfSharp.PageSize.A4;
        page.Width = _pageWidth;
        page.Height = _pageLength;
      


        _gfx = XGraphics.FromPdfPage(page);
        var rowHeight = section.PageMargin.Top;
        var pageWidth = page.Width - section.PageMargin.Left - section.PageMargin.Right - (section.Columns.Count + 1) * (section.FieldMargin.Left + section.FieldMargin.Right);
        var pageHeight = page.Height - section.PageMargin.Top - section.PageMargin.Bottom;

         //Print First Page Heading                    

        if (section.FirstPageTextFields != null)
        {
            foreach (var field in section.FirstPageTextFields)
            {
                _gfx.DrawString(string.IsNullOrEmpty(pageHeading)?field.Value: pageHeading, section.HeaderFont, field.Brush, field.Position, field.StringFormat);
            }
        }

        rowHeight += section.FirstPageHeightOffset;
        

        //Text fields
        if (section.TextFields != null)
        {
            foreach (var field in section.TextFields)
            {
                _gfx.DrawString(field.Value, field.Font, field.Brush, field.Position, field.StringFormat);
            }
        }

        //Image fields
        if (section.ImageFields != null)
        {
            foreach (var field in section.ImageFields)
            {
                _gfx.DrawImage(field.Value, field.Position);
            }
        }

        //Charts fields
        if (section.ChartFields != null)
        {
            foreach (var field in section.ChartFields)
            {
                GenerateChart(field).Draw(_gfx);
            }
        }

        return (rowHeight, pageWidth, pageHeight);
    }

    private (List<T>, string[]) SetupGrouping(List<T> rows, ReportSection section)
    {
        IOrderedEnumerable<T> orderedRows = rows.OrderBy(i => i);
        //Group by we'll be ordering by and then removing the group by column value to show grouping
        var groupByColumns = section.Columns.Where(i => i.GroupBy).ToList();
        var index = 0;
        foreach (var column in groupByColumns)
        {
            if (_properties.TryGetValue(column.DataColumn.ToLower(), out var p))
            {
                if (index == 0)
                {
                    if (column.GroupByAscending)
                    {
                        orderedRows = rows.OrderBy(o => _typeAccessor[o, p]);
                    }
                    else
                    {
                        orderedRows = rows.OrderByDescending(o => _typeAccessor[o, p]);
                    }
                }
                else
                {
                    if (column.GroupByAscending)
                    {
                        orderedRows = orderedRows.ThenBy(o => _typeAccessor[o, p]);
                    }
                    else
                    {
                        orderedRows = orderedRows.ThenByDescending(o => _typeAccessor[o, p]);
                    }
                }

                index++;
            }
        }

        List<T> displayRows = rows;

        if (groupByColumns.Any())
            displayRows = orderedRows.ToList();

        return (displayRows, new string[groupByColumns.Count]);
    }

    private void SetupColumnsAndMargins(ReportSection section, double rowHeight, double pageWidth)
    {
        var columnWidth = section.PageMargin.Left; //So we don't write at 0
        foreach (var column in section.Columns)
        {
            columnWidth += section.FieldMargin.Left; //To the left of each field
            rowHeight += section.FieldMargin.Top; //To the top of each field

            _gfx.DrawString(column.ColumnHeader, section.HeaderFont, column.Brush, new XRect(columnWidth, rowHeight, pageWidth * column.ColumnWidth / 100, section.HeaderFontHeight), column.StringFormat);
            columnWidth += pageWidth * column.ColumnWidth / section.Columns.Sum(i => i.ColumnWidth);

            columnWidth += section.FieldMargin.Right; //To the Right of each field
            rowHeight += section.FieldMargin.Bottom; //To the Bottom of each field
        }
        rowHeight += section.HeaderFontHeight;
    }

    private bool ManageDataAndRows(ReportSection section, double rowHeight, int rowIndex, double pageHeight, int totalRows, List<T> displayRows, string[] groupByValues, double pageWidth, out int currentIndex)
    {
        rowHeight += section.HeaderFontHeight;
        var displayGroupBy = true;

        while (rowHeight + section.DefaultFontHeight < pageHeight && rowIndex < totalRows)
        {
            var row = displayRows[rowIndex];
            var columnWidth = section.PageMargin.Left;
            var colIndex = 0;

            displayGroupBy = GroupRowDataFromDataObject(section, row, groupByValues, displayGroupBy);

            foreach (var column in section.Columns)
            {
                if (column.DataColumn == null)
                {
                    var filter = _columnConfigs.Where(x => x.ColumnName == column.ColumnHeader).First();
                    string value = filter.ColumnConfigFilter.ApplyFilter(row).Item4;
                    columnWidth = SetCellPlacementAndData(value, pageWidth, section, column, columnWidth, rowHeight);
                    colIndex++;
                }
                else if (_properties.TryGetValue(column.DataColumn.ToLower(), out var propName))
                {
                    string value = ConvertCellDataIntoString(propName, row);
                    value = SetGroupByValueToCol(value, displayGroupBy, column, groupByValues, colIndex);
                    columnWidth = SetCellPlacementAndData(value, pageWidth, section, column, columnWidth, rowHeight);
                    colIndex++;
                }
            }
            rowHeight += section.DefaultFontHeight;
            rowIndex++;
            displayGroupBy = false;
        }

        currentIndex = rowIndex;

        return totalRows > rowIndex;
    }

    private string SetGroupByValueToCol(string value, bool displayGroupBy, ReportColumns column, string[] groupByValues, int colIndex)
    {
        if (!displayGroupBy && column.GroupBy)
            value = "";
        else if (column.GroupBy)
        {
            groupByValues[colIndex] = value;
        }

        return value;
    }

    private double SetCellPlacementAndData(string value, double pageWidth, ReportSection section, ReportColumns column, double columnWidth, double rowHeight)
    {
        columnWidth += section.FieldMargin.Left; //To the left of each field
        rowHeight += section.FieldMargin.Top; //To the top of each field
        var fieldWidth = pageWidth * column.ColumnWidth / section.Columns.Sum(i => i.ColumnWidth);
        _gfx.DrawRectangle(XBrushes.White, new XRect(columnWidth - section.FieldMargin.Left, rowHeight - section.FieldMargin.Top, fieldWidth + section.FieldMargin.Right, section.DefaultFontHeight + section.FieldMargin.Bottom));
        _gfx.DrawString(value, section.DefaultFont, column.Brush, new XRect(columnWidth, rowHeight, fieldWidth, section.DefaultFontHeight), column.StringFormat);
        columnWidth += fieldWidth;
        columnWidth += section.FieldMargin.Right; //To the Right of each field
        rowHeight += section.FieldMargin.Bottom; //To the Bottom of each field


        return columnWidth;
    }

    private bool GroupRowDataFromDataObject(ReportSection section, T row, string[] groupByValues, bool displayGroupBy)
    {
        var colIndex = 0;

        //Figure out how far we need to write group by columns for
        foreach (var column in section.Columns.Where(i => i.GroupBy))
        {
            if (_properties.TryGetValue(column.DataColumn.ToLower(), out var propName))
            {
                var value = _typeAccessor[row, propName]?.ToString() ?? "";
                displayGroupBy = displayGroupBy || groupByValues[colIndex] != ConvertCellDataIntoString(propName, row);
                colIndex++;
            }
        }

        return displayGroupBy;
    }

    private string ConvertCellDataIntoString(string propName, T rowOfData)
    {
        dynamic value = _typeAccessor[rowOfData, propName];
        var relatedColumnConfig = _columnConfigs.Where(c => c.ClassNameMatch == propName).First();

        if (relatedColumnConfig.ColumnConfigFilter == null)
        {
            string returnValue = value?.ToString() ?? "";
            if (returnValue.Length >= relatedColumnConfig.PdfColumnTruncateWidth)
                return returnValue.Substring(0, relatedColumnConfig.PdfColumnTruncateWidth);
            return returnValue;
        }
        
        return relatedColumnConfig.ColumnConfigFilter.ApplyFilter(value).Item4;
    }
}


