using System;
namespace Excel.Component.Library.Configuration
{
    public interface IExcelPropertyConfigurationMappingFluent
    {
        IExcelPropertyConfigurationMappingFluent SetOrder(int columnOrder);
        IExcelPropertyConfigurationMappingFluent SetCaption(string caption);
    }
}
