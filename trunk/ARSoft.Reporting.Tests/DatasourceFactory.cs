namespace ARSoft.Reporting.Tests
{
    using System;
    using System.Collections.Generic;

    public class DatasourceFactory
    {
        public static TestModel GetDatasourceSimpleObject()
        {
            var datasource = new TestModel { Fecha = new DateTime(2000, 1, 2), Nombre = "pepe1", Numero = 1.5m, NumeroEntero = 1 };
            return datasource;
        }

        public static List<TestModel> GetDatasourceList()
        {
            var datasource = new List<TestModel>();
            datasource.Add(new TestModel { Fecha = new DateTime(2000, 1, 2), Nombre = "pepe1", Numero = 1 });
            datasource.Add(new TestModel { Fecha = new DateTime(2001, 2, 3), Nombre = "pepe2", Numero = 2 });
            datasource.Add(new TestModel { Fecha = new DateTime(2003, 3, 4), Nombre = "pepe3", Numero = 3 });
            return datasource;
        }
 
    }
}