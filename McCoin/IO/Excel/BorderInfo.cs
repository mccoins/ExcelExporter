using McCoin.Interfaces;

namespace McCoin.IO.Excel
{
    /// <summary>
    /// Class used to define border details
    /// </summary>
    public class BorderInfo : IBorderInfo
    {
        public string Top { get; set; }
        public string Right { get; set; }
        public string Bottom { get; set; }
        public string Left { get; set; }
        public string Width { get; set; }
    }
}
