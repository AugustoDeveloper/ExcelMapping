using System;
namespace Excel.Component.Library.Configuration
{
    public interface IExcelComplexGroupPropertyConfigurationMappingFluent : IExcelGroupPropertyConfigurationMapping
    {
        IExcelComplexGroupPropertyConfigurationMappingFluent HideHeaderPerRow(bool hideHeaderPerRow);
    }
}
