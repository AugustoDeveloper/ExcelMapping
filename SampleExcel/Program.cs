using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;
using OfficeOpenXml;
using Excel.Component.Library.Component;
using Excel.Component.Library.Component.Base;
using Excel.Component.Library.Processor;

namespace Excel.Component.Library
{
    internal class MainClass
    {
        public const string OUTPUT_PATH = @"C:\";
        public static void Main(string[] args)
        {
            var mapper = new ExcelMap();

            var ws = mapper.GetExcelWorkSheet("Content");
            
            var table = ws.CreateTable<Person>((c) =>
            {
                c.SetCaption("Teste");
                c.HideCaption(true);
            });

            table.MapProperty(p => p.Name).SetCaption("Nome");
            table.MapGroup(g =>
                {
                    g.MapProperty(p => p.Weight).SetCaption("Novo Peso");
                    g.MapProperty(p => p.Age).SetCaption("Nova Idade");
                    g.MapGroup(p =>
                    {
                        p.MapProperty(o => o.LastName).SetCaption("Sobrenome");
                        p.MapProperty(o => o.Name).SetCaption("Name");
                    }).SetCaption("Test");
                })
                .SetCaption("General Info");
            table.MapProperty(p => p.LastName).SetCaption("Sobrenome");
            table.MapProperty(p => p.Age).SetCaption("Idade");
            table.MapProperty(p => p.Weight).SetCaption("Peso");


            var processor = ExcelProcessorBuilder
                .GetDefaultBuilder()
                .WithRootContainer(table)
                .Build();

            processor.Process(ws, table,
                //new Person { Age = 10, Name = "Augusto", LastName = "Batista", Weight = 98 }
                new List<Person>{
                    new Person { Age = 10, Name = "Augqusto", LastName = "Badatista", Weight = 98 },
                    new Person { Age = 10, Name = "Augusto", LastName = "Batista", Weight = 98 },
                    new Person { Age = 10, Name = "Auguqsto", LastName = "Batiacsta", Weight = 98 },
                    new Person { Age = 10, Name = "Auguswto", LastName = "Baatista", Weight = 98 },
                    new Person { Age = 10, Name = "Augusto", LastName = "Batista", Weight = 98 },
                    new Person { Age = 10, Name = "Auguswto", LastName = "Bastacsista", Weight = 98 },
                    new Person { Age = 10, Name = "Augusto", LastName = "Batista", Weight = 98 }
                });

            var filename = "extractFile" + DateTime.Now.Ticks.ToString();
            mapper.Extract(OUTPUT_PATH, filename, true);
         
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