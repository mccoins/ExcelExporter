using System;
using System.Collections.Generic;
using McCoin.IO.Excel;
using McCoin.Interfaces;

namespace ExcelExport
{
    class Program
    {
        static void Main(string[] args)
        {        
            IExcelDoc xl = new ExcelDocument();
            List<ISource<double>> data = new List<ISource<double>>();   
            List<IColumnInfo> ColumnsList = ColumnInfoLists.ColumnsList;

            // generate simple test data
            ISource<double> td;
            for (int i = 0; i <= 5; i++)
            {
                td = new TestData(i * 2);
                data.Add(td);
            }

            // create a local excel file - testing only
            xl.CreateLocalExcelDoc(@"C:\temp\TestReport.xlsx", 
                data, ColumnsList);

            Console.WriteLine("Done");
            Console.ReadLine();
            
        }

      
    }
}
