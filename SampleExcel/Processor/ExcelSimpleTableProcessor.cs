using System.Collections;
using OfficeOpenXml;
using SampleExcel.Component.Base;

namespace SampleExcel.Processor
{
    public class ExcelSimpleTableProcessor : IExcelHeaderProcessor, IExcelRowProcessor
    {
        public int ProcessHeader(ExcelWorksheet worksheet, int row, IExcelPropertiesContainer container)
        {
            foreach (var property in container.Properties)
			{
				worksheet.Cells[row, property.ColumnOrder].Value = property.Caption;
			}

            return ++row;
        }

        public int ProcessRow(ExcelWorksheet worksheet, int row, IExcelPropertiesContainer container, IEnumerable data)
        {
			foreach (var item in data)
			{
                row = ProcessRow(worksheet, row, container, item);
			}

            return row;
        }

        public int ProcessRow(ExcelWorksheet worksheet, int row, IExcelPropertiesContainer container, object data)
        {
            foreach (var property in container.Properties)
			{
				worksheet.Cells[row, property.ColumnOrder].Value = property.GetValue(data);
			}

            return ++row;
        }
    }
}
