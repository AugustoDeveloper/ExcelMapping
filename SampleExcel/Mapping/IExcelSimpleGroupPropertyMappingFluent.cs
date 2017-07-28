using System;
using System.Linq.Expressions;
using Excel.Component.Library.Configuration;

namespace Excel.Component.Library.Mapping
{
    public interface IExcelSimpleGroupPropertyMappingFluent<TDto> : IExcelSimplePropertyMappingFluent<TDto>
    {
        IExcelSimplePropertyConfigurationMappingFluent<TDto> MapGroup(Action<IExcelSimpleGroupPropertyMappingFluent<TDto>> relationship);
        IExcelSimplePropertyConfigurationMappingFluent<TNewDto> MapGroup<TNewDto>(Expression<Func<TDto, TNewDto>> propertyExpression, Action<IExcelSimpleGroupPropertyMappingFluent<TNewDto>> relationship);
    }
}
