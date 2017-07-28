using System;
using System.Linq.Expressions;
using SampleExcel.Component.Base;
using SampleExcel.Configuration;
using SampleExcel.Mapping;

namespace SampleExcel.Component
{
    public interface IExcelSimpleProperty<TDto> : IExcelProperty
    {
        void BuildDataExtractor<TValue>(Expression<Func<TDto, TValue>> propertyExpression);
    }

    public class ExcelSimpleProperty<TDto> : IExcelSimpleProperty<TDto>, IExcelSimplePropertyConfigurationMappingFluent<TDto>, IExcelSimplePropertyMappingFluent<TDto>
    {
        private IDataExtraction _extractor;

        public int ColumnOrder { get; set; }

        public string Caption { get; set; }

        public void BuildDataExtractor<TValue>(Expression<Func<TDto, TValue>> propertyExpression)
        {
            _extractor = new DataExtraction<TDto, TValue>(propertyExpression);
        }

        public object GetValue(object data)
        {
            return _extractor.GetValueFromData(data);
        }

        public IExcelSimplePropertyConfigurationMappingFluent<TDto> Map<TValue>(Expression<Func<TDto, TValue>> propertyExpression)
        {
            BuildDataExtractor(propertyExpression);
            return this;
        }

        public IExcelPropertyConfigurationMappingFluent SetCaption(string caption)
        {
            Caption = caption;
            return this;
        }

        public IExcelPropertyConfigurationMappingFluent SetOrder(int columnOrder)
        {
            ColumnOrder = columnOrder;
            return this;
        }
    }
}
