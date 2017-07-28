using System;
using System.Collections;
using System.Collections.Generic;
using OfficeOpenXml;
using SampleExcel.Component.Base;

namespace SampleExcel.Processor
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
			var startRow = Header.ProcessHeader(worksheet, container.StartRow.Value, container);
			var enumerableData = (data as IEnumerable);
			Row.ProcessRow(worksheet, startRow, container, enumerableData == null ? data : enumerableData);
		}
    }
}
