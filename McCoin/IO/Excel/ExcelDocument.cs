using System.Collections.Generic;
using System.IO;

using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml;
using McCoin.Interfaces;

namespace McCoin.IO.Excel
{
    /// <summary>
    /// class used to build an exceldocument
    /// </summary>
    public class ExcelDocument : IExcelDoc
    {   
        /// <summary>
        /// Method used to generate an excel file and return as a MemoryStream
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="sourceList"></param>
        /// <param name="columnInfoList"></param>
        /// <returns></returns>
        public MemoryStream CreateExcelDoc<T>(List<T> sourceList, List<IColumnInfo> columnInfoList) where T : ISource<double>
        {
            IStyleSheet styleSheet = new ExcelStyleSheet(columnInfoList);
            ICell cell;

            using (MemoryStream ms = new MemoryStream())
            {
                using (SpreadsheetDocument document = SpreadsheetDocument.Create(ms, SpreadsheetDocumentType.Workbook))
                {
                    // add the workbook and work sheet
                    WorkbookPart workbookPart = document.AddWorkbookPart();
                    workbookPart.Workbook = new Workbook();

                    WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                    worksheetPart.Worksheet = new Worksheet();

                    // Adding style
                    WorkbookStylesPart stylePart = workbookPart.AddNewPart<WorkbookStylesPart>();
                    stylePart.Stylesheet = styleSheet.GenerateStyleSheet();
                    stylePart.Stylesheet.Save();

                    // Setting up columns
                    Columns columns = new Columns();

                    foreach (IColumnInfo info in columnInfoList)
                    {
                        columns.Append(new Column
                        {
                            Min = 1,
                            Max = 1,
                            Width = info.Width,
                            CustomWidth = true
                        });
                    }

                    worksheetPart.Worksheet.AppendChild(columns);

                    Sheets sheets = workbookPart.Workbook.AppendChild(new Sheets());

                    Sheet sheet = new Sheet() { Id = workbookPart.GetIdOfPart(worksheetPart), SheetId = 1, Name = "test" };

                    sheets.Append(sheet);

                    workbookPart.Workbook.Save();

                    SheetData sheetData = worksheetPart.Worksheet.AppendChild(new SheetData());
                                        
                    // Inserting each item
                    Row row;
                    foreach (var item in sourceList)
                    {
                        row = new Row();
                        cell = new ExcelCell();
                        row.Append(cell.ConstructCell(item.ExportData.ToString(), CellValues.Number));   // TODO: determine datatype based on cellInfo passed in                     
                        sheetData.AppendChild(row);
                    }

                    worksheetPart.Worksheet.Save();
                }

                return ms;
            }             
        }

        /// <summary>
        /// Method used to generate a local excel file for testing purposes
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="filename"></param>
        /// <param name="sourceList"></param>
        /// <param name="columnInfoList"></param>
        public void CreateLocalExcelDoc<T>(string fileName, List<T> sourceList, List<IColumnInfo> columnInfoList) where T : ISource<double>
        {
            ExcelStyleSheet styleSheet = new ExcelStyleSheet(columnInfoList);
            ExcelCell cell;

            using (SpreadsheetDocument document = SpreadsheetDocument.Create(fileName, SpreadsheetDocumentType.Workbook))
            {
                // add the workbook and work sheet
                WorkbookPart workbookPart = document.AddWorkbookPart();
                workbookPart.Workbook = new Workbook();

                WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                worksheetPart.Worksheet = new Worksheet();

                // create and add a style sheet to the document
                WorkbookStylesPart stylePart = workbookPart.AddNewPart<WorkbookStylesPart>();
                stylePart.Stylesheet = styleSheet.GenerateStyleSheet();
                stylePart.Stylesheet.Save();

                // Setting up columns with defined widths
                Columns columns = new Columns(); 
                    
                foreach(IColumnInfo info in columnInfoList)
                {
                    columns.Append(new Column
                    {
                        Min = 1,
                        Max = 1,
                        Width = info.Width,
                        CustomWidth = true
                    });
                }

                // add columns to the worksheet
                worksheetPart.Worksheet.AppendChild(columns);
                Sheets sheets = workbookPart.Workbook.AppendChild(new Sheets());
                Sheet sheet = new Sheet() { Id = workbookPart.GetIdOfPart(worksheetPart), SheetId = 1, Name = "test" };

                sheets.Append(sheet);

                workbookPart.Workbook.Save();                

                SheetData sheetData = worksheetPart.Worksheet.AppendChild(new SheetData());
                
                // Inserting each item based on sourceList
                Row row;
                foreach (var item in sourceList)
                {
                    row = new Row();
                    cell = new ExcelCell();                                     
                    row.Append(cell.ConstructCell(item.ExportData.ToString(), CellValues.Number));   // TODO: determine datatype based on cellInfo passed in                     
                    sheetData.AppendChild(row);
                }
                                
                worksheetPart.Worksheet.Save();
            }

        }
    }
}