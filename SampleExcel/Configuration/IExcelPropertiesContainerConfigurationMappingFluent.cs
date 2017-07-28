using System;
namespace Excel.Component.Library.Configuration
{
    public interface IExcelPropertiesContainerConfigurationMappingFluent : IExcelPropertyConfigurationMappingFluent
    {
        IExcelPropertiesContainerConfigurationMappingFluent StartAtRow(int startRow);
        IExcelPropertiesContainerConfigurationMappingFluent HideCaption(bool hideCaption);

    }
}
