﻿namespace Excel.Component.Library.Component.Base
{
    public interface IExcelProperty
    {
        IExcelPropertiesContainer Parent { get; set; }
        int ColumnOrder { get; }
        string Caption { get; }
        object GetValue(object data);
    }
}