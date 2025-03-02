using System.Reflection;
using System.Text;
using ExportWrapper.Common.Builders;
using ExportWrapper.Common.ColumnConfigFilters;
using ExportWrapper.Common.Factories;
using ExportWrapper.Common.Models;
using Microsoft.Extensions.Logging;

namespace ExportWrapper.Services.Helpers;

public static class ReportHelpers
{

    //models are connected to the main code base.. this is more of an example

    //public static byte[] SalesReportExcel(List<StoreResponse<InternalResponse<CustomReportResponse>>> storeData, List<ColumnConfig> columnConfigs, List<ExportParameters> exportParameters, ILogger log)
    //{
    //    var excelDoc = new ExcelDocumentBuilder();
    //    excelDoc.AddSheet(new ExcelSheetData().SheetData("Parameters", exportParameters, ParametersColumnConfig()));

    //    foreach (var data in storeData)
    //    {
    //        if (data.Response == null)
    //            continue;

    //        ExcelSheetData sheet = new ExcelSheetData().SheetData(data.StoreName, ConvertToSalesReportDataModel(data.Response.Data.Rows, columnConfigs, data.StoreName, log), columnConfigs);
    //        excelDoc.AddSheet(sheet);
    //    }

    //    return excelDoc.BuildDocument();
    //}

    //public static byte[] StockReportExcel(List<StoreResponse<InternalResponse<CustomReportResponse>>> storeData, List<ColumnConfig> columnConfigs, List<ExportParameters> exportParameters, ILogger log)
    //{
    //    var excelDoc = new ExcelDocumentBuilder();
    //    excelDoc.AddSheet(new ExcelSheetData().SheetData("Parameters", exportParameters, ParametersColumnConfig()));

    //    foreach (var data in storeData)
    //    {
    //        if (data.Response == null)
    //            continue;

    //        ExcelSheetData sheet = new ExcelSheetData().SheetData(data.StoreName, ConvertToStockReportDataModel(data.Response.Data.Rows, columnConfigs, data.StoreName, log), columnConfigs);
    //        excelDoc.AddSheet(sheet);
    //    }

    //    return excelDoc.BuildDocument();
    //}




    //public static byte[] SalesReportPdf(List<StoreResponse<InternalResponse<CustomReportResponse>>> storeData, List<ColumnConfig> columnConfigs, ILogger log)
    //{
    //    var pdfDoc = new PDFDocumentBuilder<SalesReportData>();
    //    Dictionary<string, List<SalesReportData>> salesReports = new Dictionary<string, List<SalesReportData>>();

    //    foreach (var data in storeData)
    //    {
    //        if (data.Response == null)
    //            continue;

    //        salesReports.Add(data.StoreName, ConvertToSalesReportDataModel(data.Response.Data.Rows, columnConfigs, data.StoreName, log));
    //    }


    //    return pdfDoc.BuildDocument(salesReports, columnConfigs, "Sales Report");
    //}

    //private static List<StockReportData> ConvertToStockReportDataModel(List<IDictionary<string, object>> reportData, List<ColumnConfig> columnConfigs, string storeName, ILogger log)
    //{
    //    List<StockReportData> stockReports = new List<StockReportData>();
    //    ;
    //    if (reportData.Count == 0)
    //    {
    //        return new List<StockReportData>();
    //    }

    //    for (int i = 0; i < reportData.Count; i++)
    //    {
    //        var stockReportRow = new StockReportData();
    //        IDictionary<string, object> rowData = reportData[i];
    //        int rowIndex = 0;
    //        stockReportRow.StoreName = storeName;

    //        foreach (var item in columnConfigs)
    //        {
    //            try
    //            {
    //                PropertyInfo piInstance = stockReportRow.GetType().GetProperty(item.ClassNameMatch);
    //                var result = rowData.TryGetValue(item.ClassNameMatch.ToString(), out object value) ? value : null;

    //                if (result == DBNull.Value)
    //                {
    //                    result = null;
    //                }
    //                piInstance.SetValue(stockReportRow, result);
    //            }
    //            catch (Exception ex)
    //            {
    //                log.LogError(ex, $"Error in ConvertToStockReportDataModel involving: {item.ColumnName}");
    //            }
    //            rowIndex++;
    //        }

    //        stockReports.Add(stockReportRow);
    //    }

    //    return stockReports;
    //}

    //private static List<SalesReportData> ConvertToSalesReportDataModel(List<IDictionary<string, object>> reportData, List<ColumnConfig> columnConfigs, string storeName, ILogger log)
    //{
    //    List<SalesReportData> salesReports = new List<SalesReportData>();
    //    ;
    //    if (reportData.Count == 0)
    //    {
    //        return new List<SalesReportData>();
    //    }

    //    for (int i = 0; i < reportData.Count; i++)
    //    {
    //        var salesReportRow = new SalesReportData();
    //        IDictionary<string, object> rowData = reportData[i];
    //        int rowIndex = 0;
    //        salesReportRow.StoreName = storeName;

    //        foreach (var item in columnConfigs)
    //        {
    //            try
    //            {
    //                PropertyInfo piInstance = salesReportRow.GetType().GetProperty(item.ClassNameMatch);
    //                var result = rowData.TryGetValue(item.ClassNameMatch, out object value) ? value : null;

    //                if (result == DBNull.Value)
    //                {
    //                    result = null;
    //                }
    //                piInstance.SetValue(salesReportRow, result);
    //            }
    //            catch (Exception ex)
    //            {
    //                log.LogError(ex, $"Error in ConvertToSalesReportDataModel involving: {item.ColumnName}");
    //            }
    //            rowIndex++;
    //        }

    //        salesReports.Add(salesReportRow);
    //    }

    //    return salesReports;
    //}

    //public static List<ColumnConfig> ParametersColumnConfig()
    //{
    //    return new List<ColumnConfig>() {
    //        new ColumnConfig("", nameof(ExportParameters.Name), 20, new BoldFilter()),
    //        new ColumnConfig("Details", nameof(ExportParameters.Value), 200)
    //    };
    //}

    //public static byte[] SalesReportCsv(List<StoreResponse<InternalResponse<CustomReportResponse>>> storeData, List<ColumnConfig> columnConfigs, List<ReportColumn> designReportColumns, ILogger log)
    //{
    //    var reportData = new List<SalesReportData>();


    //    foreach (var data in storeData)
    //    {
    //        if (data.Response == null)
    //            continue;

    //        reportData.AddRange(ConvertToSalesReportDataModel(data.Response.Data.Rows, columnConfigs, data.StoreName, log));
    //    }

    //    columnConfigs.Insert(0 , new ColumnConfig("Store Name", nameof(SalesReportData.StoreName), 20));

    //    return Encoding.ASCII.GetBytes(CSVFactory<SalesReportData>.ConvertToCsv(reportData, columnConfigs));        
    //}
}
