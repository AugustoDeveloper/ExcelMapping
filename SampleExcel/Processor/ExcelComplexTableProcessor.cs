using System.Linq;
using System.Collections;
using OfficeOpenXml;
using SampleExcel.Component.Base;
using System;
using System.Collections.Generic;

namespace SampleExcel.Processor
{
    public class ExcelComplexTableProcessor : IExcelHeaderProcessor, IExcelRowProcessor
    {
        public int ProcessHeader(ExcelWorksheet worksheet, int row, IExcelPropertiesContainer container)
        {
            var currentRow = row;
            foreach(var property in container.Properties)
            {
                var groupProperty = property as IExcelGroupProperty;
                if (groupProperty != null)
                {
                    var incrementRow = 1;
                    if (groupProperty.ShowCaption)
                    {
						worksheet.Cells[row, groupProperty.ColumnOrder, currentRow, groupProperty.GetMaxColumnOrder()].Value = groupProperty.Caption;
						worksheet.Cells[row, groupProperty.ColumnOrder, currentRow, groupProperty.GetMaxColumnOrder()].Merge = true;
                        incrementRow--;
                    }
					
                    currentRow = ProcessHeader(worksheet, currentRow + incrementRow, groupProperty) - 1;
                }
                else
                {
                    worksheet.Cells[row, property.ColumnOrder, currentRow, property.ColumnOrder].Value = property.Caption;
                    worksheet.Cells[row, property.ColumnOrder, currentRow, property.ColumnOrder].Merge = true;
                }
            }

            return ++currentRow;
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
            var complexProperties = container.Properties.OfType<IExcelComplexGroupProperty>();
            var currentRow = row;
            var maxcurrentRow = 0;


			foreach (var property in complexProperties)
            {
                var lastHeaderRow = row;
                if (property.ShowHeaderPerRow && maxcurrentRow > 0)
                {
                    lastHeaderRow = ProcessHeader(worksheet, lastHeaderRow, property);
                }

                var newData = container.GetValue(data);
                maxcurrentRow = ProcessRow(worksheet, lastHeaderRow, property, newData is IEnumerable ? (IEnumerable)newData : newData) - 1;
                currentRow = (maxcurrentRow > currentRow) ? maxcurrentRow : currentRow;
            }

			foreach (var property in container.Properties)
			{
                var groupProperty = property as IExcelComplexGroupProperty;
				if (groupProperty == null)
				{
					worksheet.Cells[row, property.ColumnOrder, currentRow, property.ColumnOrder].Value = property.GetValue(data);
					worksheet.Cells[row, property.ColumnOrder, currentRow, property.ColumnOrder].Merge = true;
				}
			}

			return ++currentRow;
        }
    }
}