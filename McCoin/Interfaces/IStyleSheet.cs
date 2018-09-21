using DocumentFormat.OpenXml.Spreadsheet;

namespace McCoin.Interfaces
{
    internal interface IStyleSheet
    {
        Stylesheet GenerateStyleSheet();
        Font LoadFont { get; }
        Fill LoadFills { get; }
        Border LoadBorder { get; }
    }
}
