using System;
using System.Collections.Generic;
using OfficeOpenXml;
using SampleExcel.Component.Base;

namespace SampleExcel.Processor
{
    public interface IExcelProcessor
    {
        void Process(ExcelWorksheet worksheet, IExcelPropertiesContainer container, object data);
    }
}
