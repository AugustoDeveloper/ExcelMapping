using System;
using Excel.Component.Library.Component.Base;
using Excel.Component.Library.Configuration;
using Excel.Component.Library.Mapping;

namespace Excel.Component.Library.Component
{
    public interface IExcelTableProperty<TDto> : IExcelTableContainer, IExcelTablePropertyMappingFluent<TDto>, IExcelTablePropertyConfigurationMappingFluent<TDto>
    {
        
    }

    public class ExcelTableProperty<TDto> : ExcelSimpleGroupProperty<TDto>, IExcelTableProperty<TDto>
    {
        public bool ShowHeaderPerRow { get; set; }

        public ExcelTableProperty(IExcelPropertiesContainer parent) : base(parent) { }

        public ExcelTableProperty() : base(null) { }

        public IExcelComplexGroupPropertyConfigurationMappingFluent HideHeaderPerRow(bool hideHeaderPerRow)
        {
            ShowHeaderPerRow = !hideHeaderPerRow;
            return this;
        }

        public IExcelTablePropertyMappingFluent<TDto> MapTable(Action<IExcelTablePropertyConfigurationMappingFluent<TDto>> relationship)
        {
            var newGroupProperty = new ExcelTableProperty<TDto>(this);
            newGroupProperty.BuildDataExtractor(data => data);
            relationship(newGroupProperty);
            Properties.Add(newGroupProperty);

            return newGroupProperty;
        }

        public IExcelTablePropertyMappingFluent<TNewDto> MapTable<TNewDto>(System.Linq.Expressions.Expression<Func<TDto, TNewDto>> propertyExpression, Action<IExcelTablePropertyConfigurationMappingFluent<TNewDto>> relationship)
        {
            var newGroupProperty = new ExcelTableProperty<TNewDto>(this);
            newGroupProperty.BuildDataExtractor(propertyExpression);
            relationship(newGroupProperty);
            Properties.Add(newGroupProperty);

            return newGroupProperty;
        }
    }
}
