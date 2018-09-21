using System.Collections.Generic;

using McCoin.IO.Excel;
using McCoin.Interfaces;

namespace ExcelExport
{
    /// <summary>
    /// Test class 
    /// </summary>
    public sealed class ColumnInfoLists
    {
        static List<IColumnInfo> _columns;
        const int COUNT = 5;

        public static List<IColumnInfo> ColumnsList
        {
            private set { }
            get
            {
                return _columns;
            }
        }

        static ColumnInfoLists()
        {
            Initialize();
        }

        private static void Initialize()
        {
            _columns = new List<IColumnInfo>();

            for (int i = 0; i < COUNT; i++)
            {
                _columns.Add(new ColumnInfo()
                {
                    DisplayOrder = 1,
                    Width = 10.0,
                    Header = new HeaderInfo
                    {
                        BGColor = "",
                        Border = new BorderInfo { Bottom = "", Left = "", Right = "", Top = "", Width = "" },
                        Font = "",
                        FontColor = "",
                        FontSize = 10,
                        HorizontalAlignment = "",
                        VerticalAlignment = ""
                    },
                    Cell = new CellInfo
                    {
                        BGColor = "",
                        Border = new BorderInfo { Bottom = "", Left = "", Right = "", Top = "", Width = "" },
                        Font = "",
                        FontColor = "",
                        FontSize = 12,
                        HorizontalAlignment = "",
                        VerticalAlignment = ""

                    }
                });
            }
        }
    }
}
