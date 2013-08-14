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
        //TODO usar directametne ExcelWriter y no ReportDedinition

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
            var streamToWrite = this.GetStreamToWrite();

            // act
            var excelWriter = CreateExcelWriter();
            var renderer = new ReportRenderer(excelWriter);
            renderer.Render(datasource, reportDefinition, streamToWrite);
            streamToWrite.Close();

            // assert
            File.Exists(this.filename).Should().Be.True();
            var excelFile = this.GetStreamToWrite();
            Executing.This(() => new HSSFWorkbook(excelFile, true)).Should().NotThrow();
        }

        [Test]
        public void RenderizaEnExcelConExpresionesEstaticasCeldasDeterminadas()
        {
            // arrange
            var reportDefinition = ReportFactory.GetReport();
            var datasource = DatasourceFactory.GetDatasourceSimpleObject();
            var streamToWrite = this.GetStreamToWrite();

            // act
            var excelWriter = CreateExcelWriter();
            var renderer = new ReportRenderer(excelWriter);
            renderer.Render(datasource, reportDefinition, streamToWrite);
            streamToWrite.Close();

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
            var streamToWrite = this.GetStreamToWrite();

            // act
            var renderer = new ReportRenderer(excelWriter);
            renderer.Render(datasource, reportDefinition, streamToWrite);
            streamToWrite.Close();

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
            var streamToWrite = this.GetStreamToWrite();

            // act
            writer.StartRender(streamToWrite);
            listContent.Write(writer, datasource);
            writer.EndRender();
            streamToWrite.Close();

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
        public void SePuedeREnderizarSobreUnaPlantilla()
        {
            // arrange
            var writer = CreateExcelWriter();
            var stream = this.GetStreamToWrite();

            // act
            var template = "template.xlt";
            writer.StartRender(stream, template);
            writer.WriteTextElement(null, null, "pepe");
            writer.EndRender();
            stream.Close();

            // assert
            var sheet = GetSheet(this.filename);
            sheet.GetRow(0).GetCell(5, MissingCellPolicy.CREATE_NULL_AS_BLANK).StringCellValue.Should().Be.EqualTo(
                "template");
        }

        private FileStream GetStreamToWrite()
        {
            return File.Open(this.filename, FileMode.OpenOrCreate, FileAccess.ReadWrite);
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