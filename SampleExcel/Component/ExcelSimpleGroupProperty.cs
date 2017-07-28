using System;
using System.Linq;
using System.Collections.Generic;
using SampleExcel.Component.Base;
using SampleExcel.Mapping;
using SampleExcel.Configuration;
using System.Linq.Expressions;

namespace SampleExcel.Component
{
    public interface IExcelSimpleGroupProperty<TDto> : IExcelGroupProperty
    {
        void BuildDataExtractor<TParentDto>(Expression<Func<TParentDto, TDto>> propertyExpression);
    }

    public class ExcelSimpleGroupProperty<TDto> : ExcelSimpleProperty<TDto>, IExcelSimpleProperty<TDto>, IExcelSimpleGroupProperty<TDto>, IExcelSimpleGroupPropertyMappingFluent<TDto>, IExcelSimpleGroupPropertyConfigurationMapping<TDto>
    {
        public List<IExcelProperty> Properties { get; private set; }

        public bool ShowCaption { get; private set; }

        public int? StartRow { get; set; }

        public int GetMaxColumnOrder()
        {
            var maxOrder = Properties.Max(p => p.ColumnOrder);
            foreach(var property in Properties.OfType<IExcelPropertiesContainer>())
            {
                var currentMaxOrder = property.GetMaxColumnOrder();
                maxOrder = maxOrder < currentMaxOrder ? currentMaxOrder : maxOrder;
            }

            return maxOrder;
        }

        public IExcelPropertiesContainerConfigurationMappingFluent HideCaption(bool hideCaption)
        {
            throw new NotImplementedException();
        }

        public IExcelSimplePropertyConfigurationMappingFluent<TDto> Map(Action<IExcelSimpleGroupPropertyMappingFluent<TDto>> relationship)
		{
            var newGroupProperty = new ExcelSimpleGroupProperty<TDto>();
            newGroupProperty.BuildDataExtractor(data => data);
			relationship(newGroupProperty);
			Properties.Add(newGroupProperty);

			return newGroupProperty;
        }

        public IExcelSimplePropertyConfigurationMappingFluent<TNewDto> Map<TNewDto>(Expression<Func<TDto, TNewDto>> propertyExpression, Action<IExcelSimpleGroupPropertyMappingFluent<TNewDto>> relationship)
        {
            var newGroupProperty = new ExcelSimpleGroupProperty<TNewDto>();
            newGroupProperty.BuildDataExtractor(propertyExpression);
            relationship(newGroupProperty);
            Properties.Add(newGroupProperty);

            return newGroupProperty;
        }

		public void BuildDataExtractor<TParentDto>(Expression<Func<TParentDto, TDto>> propertyExpression)
		{
			Extractor = new DataExtraction<TParentDto, TDto>(propertyExpression);
		}

        public IExcelPropertiesContainerConfigurationMappingFluent StartAtRow(int startRow)
        {
            StartRow = startRow;

            return this;
        }
    }
}
