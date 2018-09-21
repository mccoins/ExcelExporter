using System;
using System.Collections.Generic;
using System.IO;

namespace McCoin.Interfaces
{
    /// <summary>
    /// Interface used for the creation of an excel document
    /// </summary>
    public interface IExcelDoc
    {
        MemoryStream CreateExcelDoc<T>(List<T> sourceList, List<IColumnInfo> columnInfoList) where T : ISource<double>;

        // FOR TESTING PURPOSES ONLY
        void CreateLocalExcelDoc<T>(string fileName, List<T> sourceList, List<IColumnInfo> columnInfoList) where T : ISource<double>;  
    } 
}
