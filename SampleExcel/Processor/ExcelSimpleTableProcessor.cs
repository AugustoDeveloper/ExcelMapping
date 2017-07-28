using System.Collections;
using OfficeOpenXml;
using Excel.Component.Library.Component.Base;

namespace Excel.Component.Library.Processor
{
    public class ExcelSimpleTableProcessor : IExcelHeaderProcessor, IExcelRowProcessor
    {
        public int ProcessHeader(ExcelWorksheet worksheet, int row, IExcelPropertiesContainer container)
        {
            var incrementRow = 1;

            if (container.ShowCaption)
            {
                worksheet.Cells[row, container.ColumnOrder, row, container.GetMaxColumnOrder()].Value = container.Caption;
                worksheet.Cells[row, container.ColumnOrder, row, container.GetMaxColumnOrder()].Merge = true;
            }
            else
            {
                incrementRow--;
            }

            foreach (var property in container.Properties)
            {
                worksheet.Cells[row + incrementRow, property.ColumnOrder].Value = property.Caption;
            }

            return row++;
        }

        public int ProcessRowFromEnumerable(ExcelWorksheet worksheet, int row, IExcelPropertiesContainer container, IEnumerable data)
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
                if (property is IExcelGroupProperty)
                {
                    ProcessRow(worksheet, row, (IExcelGroupProperty)property, property.GetValue(data));
                    continue;
                }
                worksheet.Cells[row, property.ColumnOrder].Value = property.GetValue(data);
            }

            return ++row;
        }
    }
}