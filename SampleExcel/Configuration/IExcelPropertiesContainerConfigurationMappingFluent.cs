using System;
namespace SampleExcel.Configuration
{
    public interface IExcelPropertiesContainerConfigurationMappingFluent : IExcelPropertyConfigurationMappingFluent
    {
        IExcelPropertiesContainerConfigurationMappingFluent StartAtRow(int startRow);
        IExcelPropertiesContainerConfigurationMappingFluent HideCaption(bool hideCaption);

    }
}
