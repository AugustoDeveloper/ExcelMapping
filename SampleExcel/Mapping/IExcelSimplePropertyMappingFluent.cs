using System;
using System.Linq.Expressions;
using SampleExcel.Configuration;

namespace SampleExcel.Mapping
{
    public interface IExcelSimplePropertyMappingFluent<TDto>
    {
        IExcelSimplePropertyConfigurationMappingFluent<TDto> Map<TValue>(Expression<Func<TDto, TValue>> propertyExpression);
    }
}
