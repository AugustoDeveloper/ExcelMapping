﻿namespace SampleExcel.Component.Base
{
    public interface IExcelProperty
    {
        int ColumnOrder { get; }
        string Caption { get; }
        object GetValue(object data);
    }
}
