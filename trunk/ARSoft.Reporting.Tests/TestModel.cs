namespace ARSoft.Reporting.Tests
{
    using System;
    using System.Collections.Generic;

    public class TestModel
    {
        public string Nombre { get; set; }

        public DateTime Fecha { get; set; }

        public decimal Numero { get; set; }
        
        public int NumeroEntero { get; set; }

        public IEnumerable<TestModelHijo> Hijos { get; set; }
    }
}