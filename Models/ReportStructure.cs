using System.Collections.Generic;
using System.Text;
using PdfSharp.Drawing;
using PdfSharp.Charting;

namespace ZFreedom.Common.Models;

//Written by a coworker for pdf generation. I have permission to use this code outside work
public class ReportStructure
{
    public List<ReportSection> Sections { get; set; } = new List<ReportSection>();
}

public class ReportSection
{
    public List<ReportColumns> Columns { get; set; } = new List<ReportColumns>();
    public Thickness PageMargin { get; set; } = new Thickness(10, 10, 10, 10); //Here you need to take into account the Text Fields
    public Thickness FieldMargin { get; set; } = new Thickness(2, 0, 2, 0); //Here you need to take into account the Text Fields
    public List<TextField> TextFields { get; set; } = new List<TextField>();
    public List<ImageField> ImageFields { get; set; } = new List<ImageField>();
    public List<ChartField> ChartFields { get; set; } = new List<ChartField>();
    public List<TextField> FirstPageTextFields { get; set; } = new List<TextField>();
    public double FirstPageHeightOffset { get; set; } = 0; //Sometimes we have a lot of text on the first page, we want to offset the table by a bit


    public XFont HeaderFont { get; set; } = ReportDefaults.HeaderFont;
    public XFont DefaultFont { get; set; } = ReportDefaults.DefaultFont;

    public double HeaderFontHeight { get; private set; }
    public double DefaultFontHeight { get; private set; }


    //We trigger it manually inside the function, so that way we don't have to calculate it every time
    public void CalculateFontHeights()
    {
        //Calculate the font heights
        var allText = new StringBuilder();
        for (int i = 32; i < 127; i++)
        {
            allText.Append((char)i);
        }

        var defaultGraphics = XGraphics.CreateMeasureContext(new XSize(2000, 2000), XGraphicsUnit.Point, XPageDirection.Downwards);
        HeaderFontHeight = defaultGraphics.MeasureString(allText.ToString(), HeaderFont).Height;
        DefaultFontHeight = defaultGraphics.MeasureString(allText.ToString(), DefaultFont).Height;
    }
}


public class ReportOptions
{
    public List<ReportColumns> Columns { get; set; } = new List<ReportColumns>();
    public Thickness PageMargin { get; set; } = new Thickness(10, 10, 10, 10); //Here you need to take into account the Text Fields
    public Thickness FieldMargin { get; set; } = new Thickness(2, 0, 2, 0); //Here you need to take into account the Text Fields
    public List<TextField> TextFields { get; set; } = new List<TextField>();
    public List<ImageField> ImageFields { get; set; } = new List<ImageField>();
    public List<TextField> FirstPageTextFields { get; set; } = new List<TextField>();
    public double FirstPageHeightOffset { get; set; } = 0; //Sometimes we have a lot of text on the first page, we want to offset the table by a bit


    public XFont HeaderFont { get; set; } = ReportDefaults.HeaderFont;
    public XFont DefaultFont { get; set; } = ReportDefaults.DefaultFont;

    public double HeaderFontHeight { get; private set; }
    public double DefaultFontHeight { get; private set; }


    //We trigger it manually inside the function, so that way we don't have to calculate it every time
    public void CalculateFontHeights()
    {
        //Calculate the font heights
        var allText = new StringBuilder();
        for (int i = 32; i < 127; i++)
        {
            allText.Append((char)i);
        }

        var defaultGraphics = XGraphics.CreateMeasureContext(new XSize(2000, 2000), XGraphicsUnit.Point, XPageDirection.Downwards);
        HeaderFontHeight = defaultGraphics.MeasureString(allText.ToString(), HeaderFont).Height;
        DefaultFontHeight = defaultGraphics.MeasureString(allText.ToString(), DefaultFont).Height;
    }
}

public class TextField
{
    public TextField()
    {

    }

    public TextField(string value, XRect position)
    {
        Value = value;
        Position = position;
    }

    public string Value { get; set; }
    public XRect Position { get; set; }
    public XFont Font { get; set; } = ReportDefaults.DefaultFont;
    public XBrush Brush { get; set; } = XBrushes.Black;
    public XStringFormat StringFormat { get; set; } = XStringFormats.Center;
}

public class ImageField
{
    public ImageField()
    {
    }

    public ImageField(XImage value, XRect position)
    {
        Value = value;
        Position = position;
    }

    public XImage Value { get; set; }
    public XRect Position { get; set; }
}

public class ChartField
{
    public ChartType ChartType { get; set; }
    public Dictionary<string, double[]> Series { get; set; } = new Dictionary<string, double[]>();
    public string XAxisTitle { get; set; }
    public string YAxisTitle { get; set; }
    public List<string[]> XSeries { get; set; }
    public XRect Position { get; set; }
}

public class ReportColumns
{
    public ReportColumns()
    {
    }

    public ReportColumns(string columnHeader, string dataColumn, double columnWidth)
    {
        ColumnHeader = columnHeader;
        DataColumn = dataColumn;
        ColumnWidth = columnWidth;
    }

    public string ColumnHeader { get; set; }
    public string DataColumn { get; set; }
    public double ColumnWidth { get; set; }
    public XBrush Brush { get; set; } = XBrushes.Black;
    public XStringFormat StringFormat { get; set; } = XStringFormats.TopLeft;
    public bool GroupBy { get; set; } = false;
    public bool GroupByAscending { get; set; } = true;
}

public struct Thickness
{
    public Thickness(double TopBottom, double LeftRight)
    {
        Top = TopBottom;
        Left = LeftRight;
        Right = LeftRight;
        Bottom = TopBottom;
    }

    public Thickness(double left, double top, double right, double bottom)
    {
        Top = top;
        Left = left;
        Right = right;
        Bottom = bottom;
    }

    public double Top { get; set; }
    public double Left { get; set; }
    public double Right { get; set; }
    public double Bottom { get; set; }
}

public static class ReportDefaults
{

    public static XFont HeaderFont = new XFont("Calibri", 10, XFontStyleEx.Bold);
    public static XFont DefaultFont = new XFont("Calibri", 10, XFontStyleEx.Regular);

    public const double A4PortraitWidth = 595;
    public const double A4PortraitHeight = 842;
    public const double A4LandscapeWidth = A4PortraitHeight;
    public const double A4LandscapeHeight = A4PortraitWidth;
}
