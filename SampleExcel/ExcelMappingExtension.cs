using System;
using OfficeOpenXml;
using SampleExcel.Component;
using SampleExcel.Configuration;
using SampleExcel.Mapping;

namespace SampleExcel
{
    public static class ExcelMappingExtension
    {
        public static IExcelTablePropertyMappingFluent<TDto> CreateTable<TDto>(this ExcelWorksheet obj, Action<IExcelTablePropertyConfigurationMappingFluent<TDto>> configure)
        {
            var table = new ExcelTableProperty<TDto>();
            configure(table);
            return table;
        }
    }
}
