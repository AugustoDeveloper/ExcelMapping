using System;
using System.Collections.Generic;
using OfficeOpenXml;
using Excel.Component.Library.Component.Base;

namespace Excel.Component.Library.Processor
{
    public interface IExcelProcessor
    {
        void Process(ExcelWorksheet worksheet, IExcelPropertiesContainer container, object data);
    }
}
