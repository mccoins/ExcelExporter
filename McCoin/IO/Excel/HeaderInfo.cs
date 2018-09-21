using McCoin.Interfaces;

namespace McCoin.IO.Excel
{
    /// <summary>
    /// Class used to define header elements
    /// </summary>
    public class HeaderInfo : IDisplayElement
    {
        public string BGColor { get; set; }
        public string Font { get; set; }
        public string FontColor { get; set; }
        public double FontSize { get; set; }
        public string HorizontalAlignment { get; set; }
        public string VerticalAlignment { get; set; }
        public IBorderInfo Border { get; set; }
    }
}
