using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml;
using System.Reflection;
using ExportWrapper.Common.Models;
using ExportWrapper.Enums;

namespace ExportWrapper.Common.Builders;

public class ExcelDocumentBuilder
{
    private MemoryStream memory;
    private List<ExcelSheetData> sheets = new List<ExcelSheetData>();

    public ExcelDocumentBuilder()
    {
        memory = new MemoryStream();
    }

    public void AddSheet(ExcelSheetData sheetData)
    {
        sheets.Add(sheetData);
    }

    public byte[] BuildDocument()
    {
        using (SpreadsheetDocument document = SpreadsheetDocument.Create(memory, SpreadsheetDocumentType.Workbook))
        {
            WorkbookPart workbookPart = document.AddWorkbookPart();
            workbookPart.Workbook = new Workbook();

            // Add a WorkbookStylesPart
            WorkbookStylesPart workbookStylesPart = workbookPart.AddNewPart<WorkbookStylesPart>();

            // Create and define workbook styles
            workbookStylesPart.Stylesheet = GenerateStylesheet();
            workbookStylesPart.Stylesheet.Save();

            Sheets sheetsCollection = document.WorkbookPart.Workbook.AppendChild(new Sheets());

            uint sheetId = 1;

            foreach (var sheet in sheets)
            {
                WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                Sheet sheetElement = new Sheet() { Id = document.WorkbookPart.GetIdOfPart(worksheetPart), SheetId = sheetId, Name = sheet.SheetName }; //?
                sheetsCollection.Append(sheetElement);

                SheetData sheetData = new SheetData();
                Worksheet worksheet = new Worksheet();

                // Dynamically generate header based on T properties
                Row headerRow = new Row();
                Columns columns = new Columns();
                int i = 0;
                foreach (var colConfig in sheet.ColumnConfigs)
                {
                    headerRow.Append(new Cell() { CellValue = new CellValue(colConfig.ColumnName), DataType = CellValues.String, StyleIndex = 3 });
                    columns.Append(new Column()
                    {
                        BestFit = true,
                        CustomWidth = true,
                        Width = colConfig.Width,
                        Min = (uint)i + 1,
                        Max = (uint)i + 1
                    });
                    i++;
                }

                worksheet.Append(columns);
                worksheet.Append(sheetData);
                worksheetPart.Worksheet = worksheet;

                sheetData.Append(headerRow);
                var rowNumber = 0;
                foreach (dynamic item in sheet.Data)
                {
                    rowNumber++;
                    sheetData.Append(PopulateCellData(item, sheet, rowNumber));
                }
                sheetId++;
            }
   
            workbookPart.Workbook.Save();
        }

        return memory.ToArray();
    }

    private Row PopulateCellData(dynamic item, ExcelSheetData sheet, int rowNumber)
    {
        var rowStyle = GetRowConfigStyle(rowNumber, sheet);

        Row newRow = new Row();
        foreach (var conf in sheet.ColumnConfigs)
        {
            if (conf.ClassNameMatch == null)
            {
                var filteredItem = conf.ColumnConfigFilter.ApplyFilter(item, rowStyle.RowConfig);
                newRow.Append(new Cell() { CellValue = new CellValue(filteredItem.Item1), DataType = filteredItem.Item2, StyleIndex = filteredItem.Item3 == null ? rowStyle.styleId : (uint)filteredItem.Item3 });
                continue;
            }

            var prop = sheet.PropertyInfos.Where(x => x.Name == conf.ClassNameMatch).First();
            var value = prop.GetValue(item, null);

            
            if (conf.ColumnConfigFilter != null)
            {
                var filteredItem = conf.ColumnConfigFilter.ApplyFilter(value, rowStyle.RowConfig);
                newRow.Append(new Cell() { CellValue = new CellValue(filteredItem.Item1), DataType = filteredItem.Item2, StyleIndex = filteredItem.Item3 == null ? rowStyle.styleId : (uint)filteredItem.Item3 });
                continue;
            }

            CellValues dataType = CellValues.String;
            switch (prop.PropertyType.Name)
            {
                case "Char":
                case "String":
                    break;
                case "Boolean":
                    dataType = CellValues.Boolean;
                    break;
                case "List`1":
                    continue;
                case "Int32":
                    dataType = CellValues.Number;
                    int number = value == null ? 0 : (int)value;
                    newRow.Append(new Cell() { CellValue = new CellValue(number), DataType = dataType, StyleIndex = rowStyle.styleId });
                    continue;
                case "Int64":
                    dataType = CellValues.Number;
                    string largNumber = value.ToString();
                    newRow.Append(new Cell() { CellValue = new CellValue(largNumber), DataType = dataType, StyleIndex = rowStyle.styleId } );
                    continue;
                case "Nullable`1":
                    newRow.Append(ProcessNullableValues(value, prop, rowStyle.styleId, rowStyle.RowConfig));
                    continue;
                case "Double":
                    dataType = CellValues.Number;
                    double doubleValue = value == null ? 0 : (double)value;
                    newRow.Append(new Cell() { CellValue = new CellValue(doubleValue), DataType = dataType, StyleIndex = rowStyle.styleId });
                    continue;
                case "Decimal":
                    dataType = CellValues.Number;
                    decimal decimalValue = value == null ? 0 : (decimal)value;
                    newRow.Append(new Cell() { CellValue = new CellValue(decimalValue), DataType = dataType, StyleIndex = rowStyle.styleId });
                    continue;
                case "DateTime":
                    dataType = CellValues.Date;
                    string test = ((DateTime)value).ToString("dd/MM/yyyy");
                    if (((DateTime)value).ToString("dd/MM/yyyy") != "01/01/0001")
                        newRow.Append(new Cell(){ CellValue = new CellValue(((DateTime)value).ToOADate().ToString()), DataType = CellValues.Number, StyleIndex = rowStyle.styleId == 0? (uint)ExcelCellTypes.ShortDate: rowStyle.RowConfig.DateStyledCell });
                    continue;
                default:
                    break;
            }

            newRow.Append(new Cell() { CellValue = new CellValue(value), DataType = dataType });
        }
        return newRow;
    }

    private Cell ProcessNullableValues(dynamic value, PropertyInfo property, uint rowStyle, RowConfiguration rowConfiguration)
    {
        if (value == null)
            return new Cell();

        if (property.PropertyType.FullName.Contains("Int64"))
        {
            string largNumber = value == null ? "0" : value.ToString();
            return new Cell() { CellValue = new CellValue(largNumber), DataType = CellValues.Number, StyleIndex = rowStyle };
        }

        if (property.PropertyType.FullName.Contains("Boolean"))
        {
            return new Cell() { CellValue = new CellValue(value), DataType = CellValues.Boolean, StyleIndex = rowStyle };
        }

        if (property.PropertyType.FullName.Contains("DateTime"))
        {
            if (value == null)
                new Cell() { CellValue = new CellValue(), DataType = CellValues.Number, StyleIndex = (uint)ExcelCellTypes.ShortDate };

            DateTime dateTime = ((DateTime?)value).Value;
            return new Cell() { CellValue = new CellValue(dateTime.ToOADate().ToString()), DataType = CellValues.Number, StyleIndex = rowStyle == 0 ? (uint)ExcelCellTypes.ShortDate : rowConfiguration.DateStyledCell }; 
        }

        if (property.PropertyType.FullName.Contains("Int32"))
        {
            int number = value == null ? 0 : (int)value;
            return new Cell() { CellValue = new CellValue(number), DataType = CellValues.Number, StyleIndex = rowStyle };
        }

        if (property.PropertyType.FullName.Contains("Double"))
        {
            double doubleNumber = value == null ? 0 : (double)value;
            return new Cell() { CellValue = new CellValue(doubleNumber), DataType = CellValues.Number, StyleIndex = rowStyle };
        }

        if (property.PropertyType.FullName.Contains("Decimal"))
        {
            decimal doubleNumber = value == null ? 0 : (decimal)value;
            return new Cell() { CellValue = new CellValue(doubleNumber), DataType = CellValues.Number, StyleIndex = rowStyle };
        }

        return new Cell();
    }

    private Stylesheet GenerateStylesheet()
    {
        Fonts fonts = new Fonts(
            new Font(),  // Index 0 - default
            new Font( // Index 1 - sub header
                new Bold()
            ),
            new Font( // Index 2 - header
                new Bold(), new Underline()
            ));

        Fills fills = new Fills(
                new Fill(new PatternFill() { PatternType = PatternValues.None }), // Index 0 - default and can not be changed                
                new Fill(new PatternFill { PatternType = PatternValues.Gray125 }), // Index 1 - default and can not be changed
                new Fill(new PatternFill() { PatternType = PatternValues.Solid, ForegroundColor = new ForegroundColor() { Rgb = "d9dbdc" } })//foreground color is background color
            );

        Borders borders = new Borders(
            new Border // Default
            {
                LeftBorder = new LeftBorder(),
                RightBorder = new RightBorder(),
                TopBorder = new TopBorder(),
                BottomBorder = new BottomBorder(),
                DiagonalBorder = new DiagonalBorder()
            }
        );

        NumberingFormats numberFormats = new NumberingFormats(
            new NumberingFormat() { NumberFormatId = 2U, FormatCode = "\"$\"#,##0.00" });

        CellFormats cellFormats = new CellFormats(
                new CellFormat // Default section
                {
                    FontId = 0,
                    FillId = 0,
                },
                 new CellFormat
                {
                    FontId = 0,
                    FillId = 0,  
                    NumberFormatId = 2U,
                },
                new CellFormat
                {
                    FontId = 0,
                    FillId = 0,
                    NumberFormatId = 14
                },
                new CellFormat
                {
                    FontId = 1,
                    FillId = 0
                },
                new CellFormat // Grey Background section
                {
                    FontId = 0,
                    FillId = 2,
                },
                 new CellFormat
                 {
                     FontId = 0,
                     FillId = 2,
                     NumberFormatId = 2U,
                 },
                new CellFormat
                {
                    FontId = 0,
                    FillId = 2,
                    NumberFormatId = 14
                },
                new CellFormat
                {
                    FontId = 1,
                    FillId = 2
                }
      );



        var styleSheet = new Stylesheet(numberFormats, fonts, fills, borders, cellFormats) { MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "x14ac" } };
        styleSheet.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
        styleSheet.AddNamespaceDeclaration("x14ac", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac");
        return styleSheet;
    }

    public (uint styleId, RowConfiguration? RowConfig) GetRowConfigStyle(int rowNumber, ExcelSheetData sheet)
    {
        if (sheet.RowConfigurations == null)
        {
            return ((uint)ExcelCellTypes.Default, null);
        }
        var rowConfig = sheet.RowConfigurations.FirstOrDefault(x => x.RowIndex == rowNumber);

        if (rowConfig == null)
        {
            return ((uint)ExcelCellTypes.Default, null);
        }

        return ((uint)rowConfig.Color, rowConfig);
    }
}


