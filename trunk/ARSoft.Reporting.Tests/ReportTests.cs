namespace ARSoft.Reporting.Tests
{
    using System;
    using System.Collections;
    using System.Collections.Generic;
    using System.IO;
    using System.Linq;

    using NUnit.Framework;

    using SharpTestsEx;

    [TestFixture]
    public class ReportTests
    {
        [Test]
        public void ReportContieneUnaListaDeContenidos()
        {
            // arrange

            // act
            var report = this.GetReport();

            // assert
            var contents = report.Contents;
            contents.Should().Not.Be.Null();
            ((object)contents).Should().Be.InstanceOf<IEnumerable<ReportContent>>();
        }

        [Test]
        public void AlContenidoSeLePuedeAgregarUnElementoParaMostrarUnaExpresionDelModelo()
        {
            // arrange
            var report = this.GetReport();


            // act
            var expresionContent = new ExpressionContent();
            expresionContent.Expression = "model.Nombre";
            expresionContent.Position = 2; 
            report.AddContent(expresionContent);

            // assert
            report.Contents.FirstOrDefault(x => x.Equals(expresionContent)).Should().Not.Be.Null();
        }

        [Test]
        public void AlContenidoSeLePuedeAgregarUnTextoEstatico()
        {
            // arrange
            var report = this.GetReport();

            // act
            var content = new StaticContent();
            content.Text = "lalalala";
            report.AddContent(content);

            // assert
            report.Contents.FirstOrDefault(x => x.Equals(content)).Should().Not.Be.Null();
        }

        [Test]
        public void UnReportDefinitionSePuedeRenderizarComoExcel()
        {
            // arrange
            var reportDefinition = this.GetReport();

            // act
            var excelRenderer = new ExcelRenderer();
            string filename = "report.xls";
            var datasource = new List<TestModel>();
            excelRenderer.Render(datasource, reportDefinition, filename);

            // assert
            File.Exists(filename).Should().Be.True();
        }

        private ReportDefinition GetReport()
        {
            return new ReportDefinition();
        }
    }

    public class ExcelRenderer
    {
        public void Render(IEnumerable datasource, ReportDefinition reportDefinition, string filename)
        {
            throw new System.NotImplementedException();
        }
    }

    public class StaticContent : ReportContent
    {
        public string Text { get; set; }
    }

    public class ExpressionContent : ReportContent
    {
        public string Expression { get; set; }
        
        public int Position { get; set; }
    }

    public class ReportContent
    {
    }

    public class ReportDefinition
    {
        readonly IList<ReportContent> reportContents = new List<ReportContent>();

        public IEnumerable<ReportContent> Contents
        {
            get
            {
                return this.reportContents;
            }
        }

        public void AddContent(ReportContent content)
        {
            this.reportContents.Add(content);
        }
    }

    public class TestModel
    {
        public string Nombre { get; set; }

        public DateTime Fecha { get; set; }

        public decimal Numero { get; set; }

        public IEnumerable<TestModelHijo> Hijos { get; set; }
    }

    public class TestModelHijo
    {
        public string NombreHijo { get; set; }

        public DateTime FechaHijo { get; set; }

        public decimal NumeroHijo { get; set; }
    }
}