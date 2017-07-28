using System;
using System.Collections;
using System.Collections.Generic;
using OfficeOpenXml;
using Excel.Component.Library.Component.Base;

namespace Excel.Component.Library.Processor
{
    public interface IExcelRowProcessor
    {
        int ProcessRowFromEnumerable(ExcelWorksheet worksheet, int row, IExcelPropertiesContainer container, IEnumerable data);
        int ProcessRow(ExcelWorksheet worksheet, int row, IExcelPropertiesContainer container, object data);
    }
}