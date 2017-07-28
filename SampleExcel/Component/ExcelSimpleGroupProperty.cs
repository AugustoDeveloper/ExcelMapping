using System;
using System.Linq;
using System.Collections.Generic;
using Excel.Component.Library.Component.Base;
using Excel.Component.Library.Mapping;
using Excel.Component.Library.Configuration;
using System.Linq.Expressions;

namespace Excel.Component.Library.Component
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

        public ExcelSimpleGroupProperty(IExcelPropertiesContainer parent) : base(parent)
        {
            Properties = new List<IExcelProperty>();
            ShowCaption = true;
        }

        public int GetMaxColumnOrder()
        {
            var maxOrder = 0;
            if (Properties.Count > 0)
            {
                maxOrder = Properties.Max(p => p.ColumnOrder);
                foreach (var property in Properties.OfType<IExcelPropertiesContainer>())
                {
                    var currentMaxOrder = property.GetMaxColumnOrder();
                    maxOrder = maxOrder < currentMaxOrder ? currentMaxOrder : maxOrder;
                }
            }

            if (Parent != null && maxOrder == 0)
            {
                var currentMaxOrder = Parent.GetMaxColumnOrder();
                maxOrder = maxOrder < currentMaxOrder ? currentMaxOrder : maxOrder;
            }

            return maxOrder;
        }

        public IExcelPropertiesContainerConfigurationMappingFluent HideCaption(bool hideCaption)
        {
            ShowCaption = !hideCaption;
            return this;
        }

        public IExcelSimplePropertyConfigurationMappingFluent<TDto> MapGroup(Action<IExcelSimpleGroupPropertyMappingFluent<TDto>> relationship)
        {
            var newGroupProperty = new ExcelSimpleGroupProperty<TDto>(this);
            newGroupProperty.BuildDataExtractor(data => data);
            relationship(newGroupProperty);
            Properties.Add(newGroupProperty);

            return newGroupProperty;
        }

        public override IExcelSimplePropertyConfigurationMappingFluent<TDto> MapProperty<TValue>(Expression<Func<TDto, TValue>> propertyExpression)
        {
            var property = new ExcelSimpleProperty<TDto>(this);
            Properties.Add(property);
            return property.MapProperty<TValue>(propertyExpression);
        }

        public IExcelSimplePropertyConfigurationMappingFluent<TNewDto> MapGroup<TNewDto>(Expression<Func<TDto, TNewDto>> propertyExpression, Action<IExcelSimpleGroupPropertyMappingFluent<TNewDto>> relationship)
        {
            var newGroupProperty = new ExcelSimpleGroupProperty<TNewDto>(this);
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