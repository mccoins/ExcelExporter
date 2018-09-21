using System.Collections.Generic;
using System.Linq;

using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml;
using McCoin.Interfaces;

namespace McCoin.IO.Excel
{
    // TODO: Complete implementation - This is not fully functional
    /// <summary>
    /// Class used to create an excelstylesheet
    /// </summary>
    internal class ExcelStyleSheet : IStyleSheet
    {
        IColumnInfo columnInfo;
        IDisplayElement cellInfo;

        public ExcelStyleSheet(List<IColumnInfo> columnInfoList)
        {            
            columnInfo = columnInfoList.FirstOrDefault();
            cellInfo = columnInfo.Cell;
        }

        public Stylesheet GenerateStyleSheet()
        {
            return new Stylesheet(LoadFont, LoadFills, LoadBorder);
        }

        #region Private properties used to build each style element
        public Font LoadFont
        {
            get
            {
                return new Font(
                    new FontSize() { Val = cellInfo.FontSize },
                    new Bold(),
                    new Color() { Rgb = cellInfo.FontColor }); // RGB color
            }
        }
        public Fill LoadFills
        {
            get
            {
                return new Fill(
              new PatternFill(
                  new ForegroundColor { Rgb = new HexBinaryValue() { Value = cellInfo.FontColor } })
              { PatternType = PatternValues.Solid });
            }
        }
        public Border LoadBorder
        {
            // TODO: determine if border elements have been set and define them in the returning Border object
            // may consider using a border factory pattern to create the border 
            get
            {
                return new Border(
                            new LeftBorder(new Color() { Auto = true }) { Style = BorderStyleValues.Thin },
                            new RightBorder(new Color() { Auto = true }) { Style = BorderStyleValues.Thin },
                            new TopBorder(new Color() { Auto = true }) { Style = BorderStyleValues.Thin },
                            new BottomBorder(new Color() { Auto = true }) { Style = BorderStyleValues.Thin },
                            new DiagonalBorder());
            }
        }
        #endregion
    }
}
