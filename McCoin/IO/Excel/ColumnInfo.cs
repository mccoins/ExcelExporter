using McCoin.Interfaces;

namespace McCoin.IO.Excel
{
    /// <summary>
    /// Class used to define column details
    /// </summary>
    public class ColumnInfo :IColumnInfo
    {
        public int DisplayOrder { get; set; }
        public double Width { get; set; }       
        public IBorderInfo Border { get; set; }
        public enum CellType { GEMERAL, NUMBER, CURRENCY, DATE, PERCENTAGE, TEXT }
        public CellType DataType { get; set; }
        public IDisplayElement Header { get; set; }
        public IDisplayElement Cell { get; set; }        
    }

   
}
