using System.Collections.Generic;
using System.Reflection;

namespace ExportWrapper.Common.Models;
public class ExcelSheetData
{
    public string SheetName { get; set; }
    public IReadOnlyCollection<dynamic> Data { get; set; }
    public PropertyInfo[] PropertyInfos { get; set; }
    public List<ColumnConfig> ColumnConfigs { get; set; }
    public List<RowConfiguration>? RowConfigurations { get; set; }

    public ExcelSheetData SheetData<T>(string sheetName, IReadOnlyCollection<T> data, List<ColumnConfig> columnConfigs, List<RowConfiguration> rowConfigurations = null)
    {
        SheetName = sheetName.Length > 29 ? $"{sheetName.Substring(0, 29)}.." : sheetName;
        Data = data as IReadOnlyCollection<dynamic>;
        PropertyInfos = typeof(T).GetProperties();
        ColumnConfigs = columnConfigs;
        RowConfigurations = rowConfigurations;

        return this;
    }
}
