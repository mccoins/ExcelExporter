using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using McCoin.Interfaces;

namespace McCoin.IO.Excel
{
    /// <summary>
    /// class used to create an individual cells
    /// </summary>
    internal class ExcelCell : ICell
    {
        public Cell ConstructCell(string value, CellValues dataType)
        {
            return new Cell()
            {
                CellValue = new CellValue(value),
                DataType = new EnumValue<CellValues>(dataType),                
            };
        }
    }
}
