using System;
using System.Collections;
using System.Collections.Generic;
using OfficeOpenXml;
using SampleExcel.Component.Base;

namespace SampleExcel.Processor
{
    public interface IExcelRowProcessor
    {
        int ProcessRow(ExcelWorksheet worksheet, int row, IExcelPropertiesContainer container, IEnumerable data);
        int ProcessRow(ExcelWorksheet worksheet, int row, IExcelPropertiesContainer container, object data);
    }
}
