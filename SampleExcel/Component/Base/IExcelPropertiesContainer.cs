﻿using System;
using System.Collections.Generic;

namespace Excel.Component.Library.Component.Base
{
    public interface IExcelPropertiesContainer : IExcelProperty
    {
        int? StartRow { get; set; }

        bool ShowCaption { get; }

        List<IExcelProperty> Properties { get; }

        int GetMaxColumnOrder();
    }
}