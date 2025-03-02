using DocumentFormat.OpenXml.Spreadsheet;
using FastMember;
using ExportWrapper.Common.Models;
using ExportWrapper.Enums;

namespace ExportWrapper.Common.ColumnConfigFilters;

public class CalculateGPFilter<T> : IColumnConfigFilter where T : class
{
    TypeAccessor _typeAccessor = TypeAccessor.Create(typeof(T));
    private string _propName;

    public CalculateGPFilter(string sellPriceColumnName)
    {
        _propName = sellPriceColumnName;
    }

    public string[] ColumnNames { get; set; }

    public (dynamic, CellValues, uint?, string) ApplyFilter(dynamic item, RowConfiguration rowConfig)
    {
        decimal? sellPrice = _typeAccessor[item, _propName];
        decimal costPrice = item.ManagedStockCost;
        char gstCode = item.ManagedStockGstCode == null? 'Y': item.ManagedStockGstCode;
        decimal? result = CalculateGP(sellPrice, costPrice, gstCode);
        result = result.HasValue? decimal.Round(CalculateGP(_typeAccessor[item, _propName], costPrice, gstCode), 2, MidpointRounding.AwayFromZero): null;
        return (result, CellValues.Number, (uint)ExcelCellTypes.Default, result.HasValue? result.Value.ToString(): string.Empty);
    }

    public decimal? CalculateGP(decimal? sellPrice, decimal? costEx, char gstCode)
    {
        if (sellPrice == null || sellPrice == 0)
            return null;
        if (costEx == null)
            costEx = 0;

        var salesEx = sellPrice / (gstCode == 'Y' ? (decimal)1.1 : 1);
        var gp = (salesEx - costEx) / salesEx * 100;

            return gp;
    }
}
