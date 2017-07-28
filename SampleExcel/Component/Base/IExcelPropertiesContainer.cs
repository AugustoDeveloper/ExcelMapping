using System;
using System.Collections.Generic;

namespace SampleExcel.Component.Base
{
    public interface IExcelPropertiesContainer : IExcelProperty
    {
        int? StartRow { get; set; }

        List<IExcelProperty> Properties { get; }
    }
}