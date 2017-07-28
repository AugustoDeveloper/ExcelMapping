using System;
using System.Linq.Expressions;
using SampleExcel.Configuration;

namespace SampleExcel.Mapping
{
    public interface IExcelTablePropertyMappingFluent<TDto> : IExcelSimpleGroupPropertyMappingFluent<TDto>
    {
        IExcelTablePropertyMappingFluent<TDto> Map(Action<IExcelTablePropertyConfigurationMappingFluent<TDto>> relationship);
		IExcelTablePropertyMappingFluent<TNewDto> Map<TNewDto>(Expression<Func<TDto, TNewDto>> propertyExpression, Action<IExcelTablePropertyConfigurationMappingFluent<TNewDto>> relationship);
    }
}
