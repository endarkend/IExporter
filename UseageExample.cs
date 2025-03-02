using ExportWrapper.Common.Builders;
using ExportWrapper.Common.Models;

public class ExampleOfHowToUse
{

    //public async Task<byte[]> ExportSearchQuery(PriceGroupsPricingExportSearchFilterQuery query, CachedUser user)
    //{

    //    var response = (await SearchAsync(query, user)).Data.OrderByDescending(x => x.Applied).ToList();
    //    byte[] exportData;

    //    var moneyFilter = new MoneyFilter();
    //    var columnConfigs = new List<ColumnConfig>() {
    //                new ColumnConfig("Name", nameof(PriceGroupsPriceView.ManagedStockName), 40),
    //                new ColumnConfig("Cost", nameof(PriceGroupsPriceView.ManagedStockCost), 10, moneyFilter),
    //                new ColumnConfig("GP%", 7, new CalculateGPFilter<PriceGroupsPriceView>(nameof(PriceGroupsPriceView.ManagedStockBaseSellPrice))),

    //                new ColumnConfig("Base Sell Price", nameof(PriceGroupsPriceView.ManagedStockBaseSellPrice), 20, moneyFilter),
    //                new ColumnConfig("Current GP%", 17, new CalculateGPFilter<PriceGroupsPriceView>(nameof(PriceGroupsPriceView.Price))),


    //                new ColumnConfig("Current Sell Price", nameof(PriceGroupsPriceView.Price), 22, moneyFilter),
    //                new ColumnConfig("New GP%", 14, new CalculateGPFilter<PriceGroupsPriceView>(nameof(PriceGroupsPriceView.NewPrice))),

    //                new ColumnConfig("New Sell Price", nameof(PriceGroupsPriceView.NewPrice), 20, moneyFilter),
    //                new ColumnConfig("Last Applied", nameof(PriceGroupsPriceView.Applied), 35)
    //            };

    //    if (query.ExportType == Models.Enums.ExportType.csv)
    //    {
    //        exportData = Encoding.ASCII.GetBytes(CSVFactory<PriceGroupsPriceView>.ConvertToCsv(response, columnConfigs));
    //    }
    //    else if (query.ExportType == Models.Enums.ExportType.pdf)
    //    {
    //        exportData = new PDFDocumentBuilder<PriceGroupsPriceView>().BuildDocument(response, columnConfigs, "Group Pricing");
    //    }
    //    else
    //    {
    //        var excelDoc = new ExcelDocumentBuilder();
    //        ExcelSheetData sheet = new ExcelSheetData().SheetData("Group Pricing", response, columnConfigs);
    //        excelDoc.AddSheet(sheet);
    //        exportData = excelDoc.BuildDocument();
    //    }

    //    return exportData;
    //}
}