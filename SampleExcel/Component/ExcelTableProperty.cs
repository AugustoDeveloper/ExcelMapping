using System;
using SampleExcel.Component.Base;
using SampleExcel.Configuration;
using SampleExcel.Mapping;

namespace SampleExcel.Component
{
    public interface IExcelTableProperty<TDto> : IExcelTableContainer
    {
        
    }

    public class ExcelTableProperty<TDto> :ExcelSimpleGroupProperty<TDto>, IExcelTablePropertyMappingFluent<TDto>, IExcelTablePropertyConfigurationMappingFluent<TDto> ,IExcelTableProperty<TDto>
    {
        public bool ShowHeaderPerRow { get; set; }

        public IExcelComplexGroupPropertyConfigurationMappingFluent HideHedaerPerRow(bool hideHeaderPerRow)
        {
            ShowHeaderPerRow = !hideHeaderPerRow;
            return this;
        }

        public IExcelTablePropertyMappingFluent<TDto> Map(Action<IExcelTablePropertyConfigurationMappingFluent<TDto>> relationship)
        {
            var newGroupProperty = new ExcelTableProperty<TDto>();            
            newGroupProperty.BuildDataExtractor(data => data);
            relationship(newGroupProperty);
            Properties.Add(newGroupProperty);

            return newGroupProperty;
        }

        public IExcelTablePropertyMappingFluent<TNewDto> Map<TNewDto>(System.Linq.Expressions.Expression<Func<TDto, TNewDto>> propertyExpression, Action<IExcelTablePropertyConfigurationMappingFluent<TNewDto>> relationship)
        {
            
            var newGroupProperty = new ExcelTableProperty<TNewDto>();
            newGroupProperty.BuildDataExtractor(propertyExpression);
            relationship(newGroupProperty);
            Properties.Add(newGroupProperty);

            return newGroupProperty;
        }
    }
}
