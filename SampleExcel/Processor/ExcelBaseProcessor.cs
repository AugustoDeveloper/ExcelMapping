using System;
using System.Collections;
using System.Collections.Generic;
using OfficeOpenXml;
using Excel.Component.Library.Component.Base;

namespace Excel.Component.Library.Processor
{
    public class ExcelBaseProcessor : IExcelProcessor
    {
        private IExcelHeaderProcessor Header { get; set; }
        private IExcelRowProcessor Row { get; set; }

        public ExcelBaseProcessor(IExcelHeaderProcessor header ,IExcelRowProcessor row)
        {
            Header = header;
            Row = row;
        }

        public void Process(ExcelWorksheet worksheet, IExcelPropertiesContainer container, object data)
        {
            var startRow = Header.ProcessHeader(worksheet, container.StartRow.HasValue ?container.StartRow.Value : 1, container) + 1;
            var enumerableData = (data as IEnumerable);
            if (enumerableData != null)
            {
                Row.ProcessRowFromEnumerable(worksheet, startRow, container, enumerableData);
            }
            else
            {
                Row.ProcessRow(worksheet, startRow, container, data);
            }
        }
    }
}