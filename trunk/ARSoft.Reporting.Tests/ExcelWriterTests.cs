namespace ARSoft.Reporting.Tests
{
    using System.IO;
    using System.Linq;

    using NPOI.HSSF.UserModel;
    using NPOI.SS.UserModel;

    using NUnit.Framework;

    using SharpTestsEx;

    [TestFixture]
    public class ExcelWriterTests
    {
        private string filename;

        [SetUp]
        public void SetUp()
        {
            this.filename = "report.xls";
            if (File.Exists(this.filename)) File.Delete(this.filename);
        }
        
        [Test]
        public void UnReportDefinitionSePuedeRenderizarComoExcel()
        {
            // arrange
            var reportDefinition = ReportFactory.GetReport();
            var datasource = DatasourceFactory.GetDatasourceSimpleObject();

            // act
            var excelWriter = CreateExcelWriter();
            var renderer = new ReportRenderer(excelWriter);
            renderer.Render(datasource, reportDefinition, this.filename);

            // assert
            File.Exists(this.filename).Should().Be.True();
            var excelFile = new FileStream(this.filename, FileMode.Open, FileAccess.Read);
            Executing.This(() => new HSSFWorkbook(excelFile, true)).Should().NotThrow();
        }

        [Test]
        public void RenderizaEnExcelConExpresionesEstaticasCeldasDeterminadas()
        {
            // arrange
            var reportDefinition = ReportFactory.GetReport();
            var datasource = DatasourceFactory.GetDatasourceSimpleObject();

            // act
            var excelWriter = CreateExcelWriter();
            var renderer = new ReportRenderer(excelWriter);
            renderer.Render(datasource, reportDefinition, this.filename);

            // assert
            var sheet = GetSheet(this.filename);
            foreach (var content in reportDefinition.Contents.Contents.OfType<StaticContent>())
            {
                sheet.GetRow(content.X.Value).GetCell(content.Y.Value).StringCellValue.Should().Be.EqualTo(content.Text);
            }
        }

        [Test]
        public void EnExcelSiElContenidoNoTieneCoordenadasDebenIrEnColumnasContiguas()
        {
            // arrange
            var reportDefinition = ReportFactory.GetReportSinCoordenadas();
            var datasource = DatasourceFactory.GetDatasourceSimpleObject();
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
        public void ListContentEnExcelGeneraCadaItemEnUnaFila()
        {
            // arrange
            var listContent = new Reporting.ListContent();
            var nombreContent = new ExpressionContent();
            nombreContent.Expression = "model.Nombre";
            listContent.Content.AddContent(nombreContent);
            var fechaContent = new ExpressionContent();
            fechaContent.Expression = "model.Fecha";
            listContent.Content.AddContent(fechaContent);
            var writer = new ExcelWriter();
            var datasource = DatasourceFactory.GetDatasourceList();

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
    }
}