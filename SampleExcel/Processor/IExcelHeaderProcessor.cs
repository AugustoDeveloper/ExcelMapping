using System;
using OfficeOpenXml;
using SampleExcel.Component.Base;

namespace SampleExcel.Processor
{
    public interface IExcelHeaderProcessor
    {
        int ProcessHeader(ExcelWorksheet worksheet, int row, IExcelPropertiesContainer container);
    }
}
