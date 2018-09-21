using McCoin.Interfaces;

namespace McCoin.IO.Excel
{
    /// <summary>
    /// class used to define cell details for individual cells within excel columns
    /// </summary>
    public class CellInfo : IDisplayElement
    {
        public string BGColor { get; set; }
        public string Font { get; set; }
        public string FontColor { get; set; }
        public double FontSize { get; set; }
        public string HorizontalAlignment { get; set; }
        public string VerticalAlignment { get; set; }
        public IBorderInfo Border { get; set; }

        public enum CellType { GEMERAL, NUMBER, CURRENCY, DATE, PERCENTAGE, TEXT }
        public CellType DataType { get; set; }
    }
}
