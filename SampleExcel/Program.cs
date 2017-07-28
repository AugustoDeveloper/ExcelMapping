using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;
using OfficeOpenXml;
using SampleExcel.Component;
using SampleExcel.Component.Base;
using SampleExcel.Processor;

namespace SampleExcel
{
    internal class MainClass
    {
        public static void Main(string[] args)
        {
			FileInfo newFile = new FileInfo(Directory.GetCurrentDirectory() + "//sample.xlsx");
			var _package = new ExcelPackage(newFile);
            /*
            new ExcelSimpleProcessor().Process(
                _package.Workbook.Worksheets["Content"],
                1,
                new List<IExcelProperty>
            {
                (IExcelProperty) new ExcelSimpleProperty<Person>().Map(prop => prop.Name).SetCaption("Nome").SetOrder(1),
                (IExcelProperty) new ExcelSimpleProperty<Person>().Map(prop => prop.LastName).SetCaption("Sobrenome").SetOrder(2),
                (IExcelProperty) new ExcelSimpleProperty<Person>().Map(prop => prop.Age).SetCaption("Idade").SetOrder(3),
                (IExcelProperty) new ExcelSimpleProperty<Person>().Map(prop => prop.Weight).SetCaption("Peso").SetOrder(4),
            }, 
                new List<Person>
            {
                new Person { Age = 10, LastName = "Batista", Name = "Augusto", Weight = 98.0 },
                new Person { Age = 10, LastName = "Batista", Name = "Augusto", Weight = 98.30 },
                new Person { Age = 10, LastName = "Batista", Name = "Augusto", Weight = 98.20 },
            }
            );
            */
            _package.Save();
         
        }
    }



    internal class Person
    {
		public string Name { get; set; }
		public string LastName { get; set; }
		public int Age { get; set; }
		public double Weight { get; set; }
    }

}
