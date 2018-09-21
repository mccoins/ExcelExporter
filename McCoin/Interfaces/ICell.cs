using DocumentFormat.OpenXml.Spreadsheet;

namespace McCoin.Interfaces
{
    /// <summary>
    /// Interface used to describe a cell
    /// </summary>
    internal interface ICell
    {
        Cell ConstructCell(string value, CellValues dataType);
    }
}
