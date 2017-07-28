using System;
using System.Collections.Generic;
using SampleExcel.Component.Base;

namespace SampleExcel.Component
{
    public interface IExcelSimpleGroupProperty<TDto> : IExcelGroupProperty
    {
        bool ShowCaption { get; }
    }

    public class ExcelSimpleGroupProperty<TDto> : IExcelSimpleGroupProperty<TDto>
    {
        
        public List<IExcelProperty> Properties { get; private set; }

        public int ColumnOrder { get; private set; }

        public string Caption { get; private set; }

        public bool ShowCaption { get; private set; }

        public int? StartRow { get; set; }

        public object GetValue(object data)
        {
            throw new NotImplementedException();
        }
    }
}
