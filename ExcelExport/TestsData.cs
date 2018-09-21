using System;
using McCoin.Interfaces;

namespace ExcelExport
{
    /// <summary>
    /// Class used for testing purposes.
    /// </summary>
    public class TestData : ISource<double>
    {
        Random random = new Random();
        double randomNumber;

        public TestData(int min)
        {
            Initialize(min);
        }

        private void Initialize(int min)
        {
            randomNumber = random.NextDouble() * (1000 - min) + 10; 
        }

        public double ExportData
        {
            get { return randomNumber; }
        }
    }
}
