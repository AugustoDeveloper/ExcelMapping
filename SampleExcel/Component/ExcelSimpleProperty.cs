using System;
using System.Linq.Expressions;
using Excel.Component.Library.Component.Base;
using Excel.Component.Library.Configuration;
using Excel.Component.Library.Mapping;

namespace Excel.Component.Library.Component
{
    public interface IExcelSimpleProperty<TDto> : IExcelProperty
    {
        void BuildDataExtractor<TValue>(Expression<Func<TDto, TValue>> propertyExpression);
    }

    public class ExcelSimpleProperty<TDto> : IExcelSimpleProperty<TDto>, IExcelSimplePropertyConfigurationMappingFluent<TDto>, IExcelSimplePropertyMappingFluent<TDto>
    {
        protected IDataExtraction Extractor { get; set; }

        public IExcelPropertiesContainer Parent { get; set; }

        public int ColumnOrder { get; set; }

        public ExcelSimpleProperty(IExcelPropertiesContainer parent)
        {
            Parent = parent;
            ColumnOrder = Parent != null ? Parent.GetMaxColumnOrder() + 1 : 1;
        }

        public string Caption { get; set; }

        public void BuildDataExtractor<TValue>(Expression<Func<TDto, TValue>> propertyExpression)
        {
            Extractor = new DataExtraction<TDto, TValue>(propertyExpression);
        }

        public object GetValue(object data)
        {
            return Extractor.GetValueFromData(data);
        }

        public virtual IExcelSimplePropertyConfigurationMappingFluent<TDto> MapProperty<TValue>(Expression<Func<TDto, TValue>> propertyExpression)
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