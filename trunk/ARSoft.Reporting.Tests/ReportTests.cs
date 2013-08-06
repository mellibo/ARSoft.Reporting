namespace ARSoft.Reporting.Tests
{
    using System;
    using System.Collections;
    using System.Collections.Generic;
    using System.IO;
    using System.Linq;

    using NPOI.HSSF.UserModel;
    using NPOI.SS.UserModel;

    using NUnit.Framework;

    using SharpTestsEx;

    [TestFixture]
    public class ReportTests
    {
        private string filename;

        [SetUp]
        public void SetUp()
        {
            filename = "report.xls";
            if (File.Exists(this.filename)) File.Delete(this.filename);
        }
        
            [Test]
        public void ReportContieneUnaListaDeContenidos()
        {
            // arrange

            // act
            var report = this.GetReport();

            // assert
            var contents = report.Contents.Contents as object;
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
            expresionContent.Y = 2; 
            report.Contents.AddContent(expresionContent);

            // assert
            report.Contents.Contents.FirstOrDefault(x => x.Equals(expresionContent)).Should().Not.Be.Null();
        }

        [Test]
        public void AlContenidoSeLePuedeAgregarUnTextoEstatico()
        {
            // arrange
            var report = this.GetReport();

            // act
            var content = new StaticContent();
            content.Text = "lalalala";
            report.Contents.AddContent(content);

            // assert
            report.Contents.Contents.FirstOrDefault(x => x.Equals(content)).Should().Not.Be.Null();
        }

        [Test]
        public void UnReportDefinitionSePuedeRenderizarComoExcel()
        {
            // arrange
            var reportDefinition = this.GetReport();
            var datasource = GetDatasourceSimpleObject();

            // act
            var excelWriter = CreateExcelWriter();
            var renderer = new ReportRenderer(excelWriter);
            renderer.Render(datasource, reportDefinition, this.filename);

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
            var datasource = GetDatasourceSimpleObject();

            // act
            var excelWriter = CreateExcelWriter();
            var renderer = new ReportRenderer(excelWriter);
            renderer.Render(datasource, reportDefinition, filename);

            // assert
            var sheet = GetSheet(filename);
            foreach (var content in reportDefinition.Contents.Contents.OfType<StaticContent>())
            {
                sheet.GetRow(content.X.Value).GetCell(content.Y.Value).StringCellValue.Should().Be.EqualTo(content.Text);
            }
        }

        [Test]
        public void RenderizarUnaExpresionEvaluadaSobreElModeloEnUnaPosicionDeterminada()
        {
            // arrange
            var reportDefinition = this.GetReport();
            var datasource = GetDatasourceSimpleObject();

            // act
            var excelWriter = CreateExcelWriter();
            var renderer = new ReportRenderer(excelWriter);
            renderer.Render(datasource, reportDefinition, filename);

            // assert
            var sheet = GetSheet(filename);

            var content =
                reportDefinition.Contents.Contents.OfType<ExpressionContent>().FirstOrDefault(
                    x => x.Expression == "model.NumeroEntero");
            sheet.GetRow(content.X.Value).GetCell(content.Y.Value).StringCellValue.Should().Be.EqualTo(datasource.NumeroEntero.ToString());
        }

        [Test]
        public void ListContentTieneUnaExpresionParaObtenerElListadoDesdeElModeloYenElWriteDebeIterarLaLista()
        {
            // arrange
            var listContent = new ListContent();
            listContent.Content.AddContent(new StaticContent());
            var writer = new MockWriter();
            writer.StartRender("nada");

            // act
            //listContent.DataSourceExpression = "model"; // el modelo es la propia lista
            var datasourceList = GetDatasourceList() as IList;
            listContent.Write(writer, datasourceList);

            // assert
            writer.WriteCount.Should().Be.EqualTo(datasourceList.Count);
        }

        [Test]
        public void ListContentEnExcelGeneraCadaItemEnUnaFila()
        {
            // arrange
            var listContent = new ListContent();
            var nombreContent = new ExpressionContent();
            nombreContent.Expression = "model.Nombre";
            listContent.Content.AddContent(nombreContent);
            var fechaContent = new ExpressionContent();
            fechaContent.Expression = "model.Fecha";
            listContent.Content.AddContent(fechaContent);
            var writer = new ExcelWriter();
            var datasource = GetDatasourceList();

            // act
            writer.StartRender(this.filename);
            listContent.Write(writer, datasource);
            writer.EndRender();

            // assert
            var sheet = GetSheet(this.filename);
            var i = 0;
            foreach (var item in datasource)
            {
                var row = sheet.GetRow(i);
                row.Should("cantidad de rows = items").Not.Be.Null();
                row.GetCell(0).StringCellValue.Should().Be.EqualTo(item.Nombre);
                row.GetCell(1).StringCellValue.Should().Be.EqualTo(item.Fecha.ToString());
                i++;
            }
        }

        [Test]
        public void EnExcelSiElContenidoNoTieneCoordenadasDebenIrEnColumnasContiguas()
        {
            // arrange
            var reportDefinition = this.GetReportSinCoordenadas();
            var datasource = GetDatasourceSimpleObject();
            var excelWriter = CreateExcelWriter();

            // act
            var renderer = new ReportRenderer(excelWriter);
            renderer.Render(datasource, reportDefinition, this.filename);

            // assert
            var sheet = GetSheet(this.filename);

            var row = sheet.GetRow(0);
            for (int i = 0; i < reportDefinition.Contents.Contents.Count(); i++)
            {
                row.Cells[i].ColumnIndex.Should().Be.EqualTo(i);
            }
        }

        [Test]
        public void ListContentCreaUnaNuevaFilaPorCadaItem()
        {
            // arrange


            // act


            // assert
            Assert.Inconclusive();
        }


        private static ISheet GetSheet(string filename)
        {
            var excelFile = new FileStream(filename, FileMode.Open, FileAccess.Read);
            var workbook = new HSSFWorkbook(excelFile, true);
            var sheet = workbook.GetSheet("hoja1");
            return sheet;
        }

        private static ExcelWriter CreateExcelWriter()
        {
            var excelWriter = new ExcelWriter();
            return excelWriter;
        }

        private static TestModel GetDatasourceSimpleObject()
        {
            var datasource = new TestModel
                { Fecha = new DateTime(2000, 1, 2), Nombre = "pepe1", Numero = 1.5m, NumeroEntero = 1 };
            return datasource;
        }

        private static List<TestModel> GetDatasourceList()
        {
            var datasource = new List<TestModel>();
            datasource.Add(new TestModel { Fecha = new DateTime(2000, 1, 2), Nombre = "pepe1", Numero = 1 });
            datasource.Add(new TestModel { Fecha = new DateTime(2001, 2, 3), Nombre = "pepe2", Numero = 2 });
            datasource.Add(new TestModel { Fecha = new DateTime(2003, 3, 4), Nombre = "pepe3", Numero = 3 });
            return datasource;
        }

        private ReportDefinition GetReport()
        {
            var reportDefinition = new ReportDefinition();
            var staticContent = new StaticContent();
            staticContent.Text = "pepe";
            staticContent.X = 5;
            staticContent.Y = 6;
            reportDefinition.Contents.AddContent(staticContent);

            staticContent = new StaticContent();
            staticContent.Text = "pepe 2";
            staticContent.X = 7;
            staticContent.Y = 3;
            reportDefinition.Contents.AddContent(staticContent);

            var expressionContent = new ExpressionContent();
            expressionContent.Expression = "model.NumeroEntero";
            expressionContent.Y = 2;
            expressionContent.X = 4;
            reportDefinition.Contents.AddContent(expressionContent);
            return reportDefinition;
        }

        private ReportDefinition GetReportSinCoordenadas()
        {
            var reportDefinition = new ReportDefinition();
            var staticContent = new StaticContent();
            staticContent.Text = "pepe";
            reportDefinition.Contents.AddContent(staticContent);

            staticContent = new StaticContent();
            staticContent.Text = "pepe 2";
            reportDefinition.Contents.AddContent(staticContent);

            staticContent = new StaticContent();
            staticContent.Text = "pepe 3";
            reportDefinition.Contents.AddContent(staticContent);

            return reportDefinition;
        }

    }

    public class MockWriter : IReportWriter
    {
        public int WriteCount { get; set; }

        public void StartRender(string filename)
        {
            WriteCount = 0;
        }

        public void EndRender()
        {
            
        }

        public void WriteTextElement(int? x, int? y, string text)
        {
            WriteCount++;
        }

        public void NewRow()
        {
            
        }
    }

    public class TestModel
    {
        public string Nombre { get; set; }

        public DateTime Fecha { get; set; }

        public decimal Numero { get; set; }
        
        public int NumeroEntero { get; set; }

        public IEnumerable<TestModelHijo> Hijos { get; set; }
    }

    public class TestModelHijo
    {
        public string NombreHijo { get; set; }

        public DateTime FechaHijo { get; set; }

        public decimal NumeroHijo { get; set; }
    }
}