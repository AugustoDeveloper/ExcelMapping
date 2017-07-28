using System;
using OfficeOpenXml;
using Excel.Component.Library.Component;
using Excel.Component.Library.Configuration;
using Excel.Component.Library.Mapping;

namespace Excel.Component.Library
{
    public static class ExcelMappingExtension
    {
        public static IExcelTableProperty<TDto> CreateTable<TDto>(this ExcelWorksheet obj, Action<IExcelTablePropertyConfigurationMappingFluent<TDto>> configure)
        {
            var table = new ExcelTableProperty<TDto>();
            configure(table);
            return table;
        }
    }
}