
using System.Reflection;
using System.Text;
using ExportWrapper.Common.Models;


namespace ExportWrapper.Common.Factories;
public static class CSVFactory<T> where T : class
{
    public static string ConvertToCsv(IReadOnlyCollection<T> data, IReadOnlyCollection<ColumnConfig> columnConfigs)
    {
        List<PropertyInfo> properties = new List<PropertyInfo>();
        var nonFilterOnlyProperties = columnConfigs.Where(p => p.ClassNameMatch != null).ToList();

        foreach (var co in nonFilterOnlyProperties)
        {
            properties.Add(typeof(T).GetProperties().Where(x => co.ClassNameMatch == x.Name).First());
        };

        return BuildCsvData(properties, data, columnConfigs);
    }

    private static string BuildCsvData(IReadOnlyCollection<PropertyInfo> properties, IReadOnlyCollection<T> data, IReadOnlyCollection<ColumnConfig> columnConfigs)
    {
        StringBuilder csvBuilder = new StringBuilder();
        AppendHeader(csvBuilder, columnConfigs);
        AppendDataRows(csvBuilder, properties, data, columnConfigs);
        return csvBuilder.ToString();
    }

    private static void AppendHeader(StringBuilder csvBuilder, IReadOnlyCollection<ColumnConfig> columnConfigs)
    {
        foreach (var conf in columnConfigs)
        {
            csvBuilder.Append(conf.ColumnName + ",");
        }
        csvBuilder.Remove(csvBuilder.Length - 1, 1);
        csvBuilder.AppendLine();
    }

    private static void AppendDataRows(StringBuilder csvBuilder, IReadOnlyCollection<PropertyInfo> properties, IReadOnlyCollection<T> data, IReadOnlyCollection<ColumnConfig> columnConfigs)
    {
        foreach (T item in data)
        {
            foreach (var config in columnConfigs)
            {
                if (config.ClassNameMatch == null)
                {
                    string valueString = config.ColumnConfigFilter.ApplyFilter(item).Item4;
                    csvBuilder.Append(valueString + ",");
                    continue;
                }
                var property = properties.Where(x => x.Name == config.ClassNameMatch).First();
                if (config.ColumnConfigFilter == null)
                {
                    object value = property.GetValue(item, null);
                    string valueString = value?.ToString().Replace(",", ";") ?? String.Empty;
                    csvBuilder.Append(valueString + ",");
                }                
                else
                {
                    string valueString = config.ColumnConfigFilter.ApplyFilter(property.GetValue(item, null)).Item4.Replace(",", ";") ?? String.Empty;
                    csvBuilder.Append(valueString + ",");
                }
            }
            csvBuilder.Remove(csvBuilder.Length - 1, 1);
            csvBuilder.AppendLine();
        }
    }
}
