using System;
namespace SampleExcel.Configuration
{
    public interface IExcelPropertyConfigurationMappingFluent
    {
        IExcelPropertyConfigurationMappingFluent SetOrder(int columnOrder);
        IExcelPropertyConfigurationMappingFluent SetCaption(string caption);
    }
}
