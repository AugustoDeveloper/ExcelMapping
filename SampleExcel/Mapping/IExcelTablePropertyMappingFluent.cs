using System;
using System.Linq.Expressions;
using Excel.Component.Library.Configuration;

namespace Excel.Component.Library.Mapping
{
    public interface IExcelTablePropertyMappingFluent<TDto> : IExcelSimpleGroupPropertyMappingFluent<TDto>
    {
        IExcelTablePropertyMappingFluent<TDto> MapTable(Action<IExcelTablePropertyConfigurationMappingFluent<TDto>> relationship);
		IExcelTablePropertyMappingFluent<TNewDto> MapTable<TNewDto>(Expression<Func<TDto, TNewDto>> propertyExpression, Action<IExcelTablePropertyConfigurationMappingFluent<TNewDto>> relationship);
    }
}
