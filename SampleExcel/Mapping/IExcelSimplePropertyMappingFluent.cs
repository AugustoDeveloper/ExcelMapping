using System;
using System.Linq.Expressions;
using Excel.Component.Library.Configuration;

namespace Excel.Component.Library.Mapping
{
    public interface IExcelSimplePropertyMappingFluent<TDto>
    {
        IExcelSimplePropertyConfigurationMappingFluent<TDto> MapProperty<TValue>(Expression<Func<TDto, TValue>> propertyExpression);
    }
}
