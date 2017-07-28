using System;
using System.Collections;
using OfficeOpenXml;
using SampleExcel.Component.Base;

namespace SampleExcel.Processor
{
    public class ExcelComplexTableProcessor : IExcelHeaderProcessor, IExcelRowProcessor
    {
        public void Process(ExcelWorksheet worksheet, IExcelPropertiesContainer container, object data)
        {
            throw new NotImplementedException();
        }

        public int ProcessHeader(ExcelWorksheet worksheet, int row, IExcelPropertiesContainer container)
        {
            throw new NotImplementedException();
        }

        public int ProcessRow(ExcelWorksheet worksheet, int row, IExcelPropertiesContainer container, IEnumerable data)
        {
            throw new NotImplementedException();
        }

        public int ProcessRow(ExcelWorksheet worksheet, int row, IExcelPropertiesContainer container, object data)
        {
            throw new NotImplementedException();
        }
    }
}