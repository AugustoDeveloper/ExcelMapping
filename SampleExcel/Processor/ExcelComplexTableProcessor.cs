using System.Linq;
using System.Collections;
using OfficeOpenXml;
using Excel.Component.Library.Component.Base;
using System;
using System.Collections.Generic;

namespace Excel.Component.Library.Processor
{
    public class ExcelComplexTableProcessor : IExcelHeaderProcessor, IExcelRowProcessor
    {
        public int ProcessHeader(ExcelWorksheet worksheet, int row, IExcelPropertiesContainer container)
        {
            var currentRow = row;

            var incrementRow = 1;

            var groupProperties = container.Properties.OfType<IExcelGroupProperty>();
            var maxcurrentRow = 0;

            foreach (var property in groupProperties)
            {
                var lastHeaderRow = row;

                maxcurrentRow = ProcessHeader(worksheet, lastHeaderRow + 1, property);
                currentRow = (maxcurrentRow > currentRow) ? maxcurrentRow : currentRow;
            }

            if (container.ShowCaption)
            {
                worksheet.Cells[row - 1, container.ColumnOrder, row - 1, container.GetMaxColumnOrder()].Value = container.Caption;
                worksheet.Cells[row - 1, container.ColumnOrder, row - 1, container.GetMaxColumnOrder()].Merge = true;
            }
            else
            {
                incrementRow--;
            }

            foreach(var property in container.Properties)
            {
                var groupProperty = property as IExcelGroupProperty;

                if (groupProperty == null)
                {
                    worksheet.Cells[row, property.ColumnOrder, currentRow, property.ColumnOrder].Value = property.Caption;
                    worksheet.Cells[row, property.ColumnOrder, currentRow, property.ColumnOrder].Merge = true;
                }
            }

            return currentRow;
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
                if (newData is IEnumerable)
                {
                    maxcurrentRow = ProcessRowFromEnumerable(worksheet, lastHeaderRow, property, (IEnumerable)newData) - 1;
                }
                else
                {
                    maxcurrentRow = ProcessRow(worksheet, lastHeaderRow, property, newData) - 1;
                }
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