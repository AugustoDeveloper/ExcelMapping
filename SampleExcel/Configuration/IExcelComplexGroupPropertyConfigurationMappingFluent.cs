using System;
namespace SampleExcel.Configuration
{
    public interface IExcelComplexGroupPropertyConfigurationMappingFluent : IExcelGroupPropertyConfigurationMapping
    {
        IExcelComplexGroupPropertyConfigurationMappingFluent HideHedaerPerRow(bool hideHeaderPerRow);
    }
}
