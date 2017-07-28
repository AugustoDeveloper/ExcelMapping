using System;
using System.Linq;
using System.Collections.Generic;
using System.Diagnostics.Contracts;
using SampleExcel.Component.Base;

namespace SampleExcel.Processor
{
    public interface IExcelProcessorBuilder
    {
        IExcelProcessorBuilder WithRootContainer(IExcelPropertiesContainer container);
        IExcelProcessor Build();
    }

    public class ExcelProcessorBuilder : IExcelProcessorBuilder
    {
		private IExcelHeaderProcessor _header;
		private IExcelRowProcessor _row;

        private ExcelProcessorBuilder() { }

        public IExcelProcessor Build()
        {
            return new ExcelBaseProcessor(_header, _row);
        }

        public IExcelProcessorBuilder WithRootContainer(IExcelPropertiesContainer container)
        {
            _header = BuildHeader(container.Properties);
            _row = BuildRow(container.Properties);
            return this;
        }

        private IExcelHeaderProcessor BuildHeader(List<IExcelProperty> properties)
		{
            var hasGroupProperty = properties.Count(p => p is IExcelGroupProperty) > 0;
			if (hasGroupProperty)
				return new ExcelComplexTableProcessor();
			else
				return new ExcelSimpleTableProcessor();
		}

		private IExcelRowProcessor BuildRow(List<IExcelProperty> properties)
		{
			var hasComplexGroupProperty = properties.Count(p => p is IExcelComplexGroupProperty) > 0;
            if (hasComplexGroupProperty)
                return new ExcelComplexTableProcessor();
            else
                return new ExcelSimpleTableProcessor();
		}

        public static IExcelProcessorBuilder GetDefaultBuilder()
        {
            return new ExcelProcessorBuilder();
        }
    }
}