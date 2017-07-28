using System;
using OfficeOpenXml;
using Excel.Component.Library.Component.Base;

namespace Excel.Component.Library.Processor
{
    public interface IExcelHeaderProcessor
    {
        int ProcessHeader(ExcelWorksheet worksheet, int row, IExcelPropertiesContainer container);
    }
}
