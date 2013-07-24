namespace ARSoft.Reporting.Tests
{
    using System;
    using System.Collections.Generic;
    using System.IO;
    using System.Linq;

    using NPOI.HSSF.UserModel;

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
            var contents = report.Contents as object;
            contents.Should().Not.Be.Null();
            contents.Should().Be.InstanceOf<IEnumerable<ReportContent>>();
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
            using (var st = File.Create(filename))
            {
                excelRenderer.Render(datasource, reportDefinition, st);
                st.Close();
            }

            // assert
            File.Exists(filename).Should().Be.True();
            var excelFile = new FileStream(filename, FileMode.Open, FileAccess.Read);
            HSSFWorkbook workbook;
            Executing.This(() => workbook = new HSSFWorkbook(excelFile, true)).Should().NotThrow();
        }

        [Test]
        public void RenderizaEnExcelConExpresionesEstaticasCeldasDeterminadas()
        {
            // arrange
            var reportDefinition = this.GetReport();

            // act
            var excelRenderer = new ExcelRenderer();
            string filename = "report.xls";
            var datasource = new List<TestModel>();
            using (var st = File.Create(filename))
            {
                excelRenderer.Render(datasource, reportDefinition, st);
                st.Close();
            }

            // assert
            var excelFile = new FileStream(filename, FileMode.Open, FileAccess.Read);
            var workbook = new HSSFWorkbook(excelFile, true);
            var sheet = workbook.GetSheet("hoja1");
            foreach (var content in reportDefinition.Contents.OfType<StaticContent>())
            {
                sheet.GetRow(content.X.Value - 1).GetCell(content.Y.Value - 1).StringCellValue.Should().Be.EqualTo(content.Text);
            }
        }

        private ReportDefinition GetReport()
        {
            var reportDefinition = new ReportDefinition();
            var staticContent = new StaticContent();
            staticContent.Text = "pepe";
            staticContent.X = 5;
            staticContent.Y = 6;

            staticContent = new StaticContent();
            staticContent.Text = "pepe 2";
            staticContent.X = 7;
            staticContent.Y = 3;

            reportDefinition.AddContent(staticContent);
            return reportDefinition;
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