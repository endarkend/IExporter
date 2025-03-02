

using ExportWrapper.Enums;

namespace ExportWrapper.Common.Models;
public class RowConfiguration
{
    public RowConfiguration(int rowIndex, ExcelRowColors color) {
        RowIndex = rowIndex;
        Color = color;
    }

    public int RowIndex { get; set; }

    public ExcelRowColors Color { get; set; }
    
    public uint BoldStyledCell { get { return (uint)Color + 3; } }

    public uint DateStyledCell { get { return (uint)Color + 2; } }

    public uint MoneyStyledCell { get { return (uint)Color + 1; } }
}
