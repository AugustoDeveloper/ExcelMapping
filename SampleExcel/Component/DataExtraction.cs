using System;
using System.Linq.Expressions;

namespace SampleExcel.Component
{
	public interface IDataExtraction
	{
		object GetValueFromData(object data);
	}

	public class DataExtraction<TDto, TValue> : IDataExtraction
	{
		private Func<TDto, TValue> _compiledFunction;

		public DataExtraction(Expression<Func<TDto, TValue>> propertyExpression)
		{
			_compiledFunction = propertyExpression.Compile();
		}

		public object GetValueFromData(object data)
		{
			return _compiledFunction.Invoke((TDto)data);
		}
	}
}
