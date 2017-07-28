using System;
using System.Linq.Expressions;
using SampleExcel.Configuration;

namespace SampleExcel.Mapping
{
    public interface IExcelSimpleGroupPropertyMappingFluent<TDto> : IExcelSimplePropertyMappingFluent<TDto>
    {
        IExcelSimplePropertyConfigurationMappingFluent<TDto> Map(Action<IExcelSimpleGroupPropertyMappingFluent<TDto>> relationship);
        IExcelSimplePropertyConfigurationMappingFluent<TNewDto> Map<TNewDto>(Expression<Func<TDto, TNewDto>> propertyExpression, Action<IExcelSimpleGroupPropertyMappingFluent<TNewDto>> relationship);
    }
}
