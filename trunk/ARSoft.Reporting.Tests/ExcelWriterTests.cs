namespace ARSoft.Reporting.Tests
{
    using System.IO;
    using System.Linq;

    using global::NPOI.HSSF.UserModel;

    using NUnit.Framework;

    using SharpTestsEx;

    using NPOIUserModel = global::NPOI.SS.UserModel;

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
                sheet.GetRow(content.Y.Value).GetCell(content.X.Value).StringCellValue.Should().Be.EqualTo(content.Text);
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
                row.Cells[i].ColumnIndex.Should("cell" + i).Be.EqualTo(i);
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
            var writer = WriterFactory.ExcelWriter();
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
            sheet.GetRow(0).GetCell(5, NPOIUserModel.MissingCellPolicy.CREATE_NULL_AS_BLANK).StringCellValue.Should().Be.EqualTo(
                "template");
        }

        [Test]
        public void ListContentEmpiezaARenderizarDesdeUnaPosicionXY()
        {
            // arrange
            var listContent = ListContentWithRowNumber();
            var stream = this.GetStreamToWrite();
            var writer = CreateExcelWriter();
            var template = "template.xlt";
            writer.StartRender(stream, template);

            // act
            listContent.X = 5;
            listContent.Y = 2;
            listContent.Write(writer, DatasourceFactory.GetDatasourceList());

            writer.EndRender();
            stream.Close();

            // assert
            var sheet = GetSheet(this.filename);
            sheet.GetRow(listContent.Y.Value).GetCell(listContent.X.Value, NPOIUserModel.MissingCellPolicy.CREATE_NULL_AS_BLANK).StringCellValue.Should().Be.EqualTo(
                "1");
            sheet.GetRow(listContent.Y.Value + 1).GetCell(listContent.X.Value, NPOIUserModel.MissingCellPolicy.CREATE_NULL_AS_BLANK).StringCellValue.Should().Be.EqualTo(
                "2");
            sheet.GetRow(listContent.Y.Value + 2).GetCell(listContent.X.Value, NPOIUserModel.MissingCellPolicy.CREATE_NULL_AS_BLANK).StringCellValue.Should().Be.EqualTo(
                "3");
        }

        [Test]
        public void ListContentPuedeSuministrarUnItemTemplate()
        {
            // arrange
            var listContent = ListContentWithRowNumber();
            var stream = this.GetStreamToWrite();
            var writer = CreateExcelWriter();
            var template = "template.xlt";
            writer.StartRender(stream, template);

            // act
            listContent.X = 5;
            listContent.Y = 2;
            listContent.ItemTemplates.Add(typeof(ExcelWriter), "15");
            listContent.Write(writer, DatasourceFactory.GetDatasourceList());

            writer.EndRender();
            stream.Close();

            // assert
            var sheet = GetSheet(this.filename);
            for (int i = 0; i < DatasourceFactory.GetDatasourceList().Count; i++)
            {
                sheet.GetRow(listContent.Y.Value + i).GetCell(listContent.X.Value + 1, NPOIUserModel.MissingCellPolicy.CREATE_NULL_AS_BLANK).NumericCellValue.Should().Be.EqualTo(
                    2);
                sheet.GetRow(listContent.Y.Value + i).GetCell(listContent.X.Value + 2, NPOIUserModel.MissingCellPolicy.CREATE_NULL_AS_BLANK).NumericCellValue.Should().Be.EqualTo(
                    3);
                sheet.GetRow(listContent.Y.Value + i).GetCell(listContent.X.Value + 3, NPOIUserModel.MissingCellPolicy.CREATE_NULL_AS_BLANK).NumericCellValue.Should().Be.EqualTo(
                    4);
                sheet.GetRow(listContent.Y.Value + i).GetCell(listContent.X.Value + 4, NPOIUserModel.MissingCellPolicy.CREATE_NULL_AS_BLANK).NumericCellValue.Should().Be.EqualTo(
                    5);
                for (var j = 0; j < 5; j++)
                {
                    sheet.GetRow(listContent.Y.Value + i).GetCell(
                        listContent.X.Value + j, NPOIUserModel.MissingCellPolicy.CREATE_NULL_AS_BLANK).CellStyle.Alignment.
                        Should().Be.EqualTo(NPOIUserModel.HorizontalAlignment.CENTER);
                    sheet.GetRow(listContent.Y.Value + i).GetCell(
                        listContent.X.Value + j, NPOIUserModel.MissingCellPolicy.CREATE_NULL_AS_BLANK).CellStyle.
                        BorderBottom.Should().Be.EqualTo(NPOIUserModel.BorderStyle.THIN);
                    sheet.GetRow(listContent.Y.Value + i).GetCell(
                        listContent.X.Value + j, NPOIUserModel.MissingCellPolicy.CREATE_NULL_AS_BLANK).CellStyle.BorderTop.
                        Should().Be.EqualTo(NPOIUserModel.BorderStyle.THIN);
                }
            }
        }
        private static ListContent ListContentWithRowNumber()
        {
            var listContent = new ListContent();
            listContent.Content.AddContent(new ExpressionContent { Expression = "Context.ItemNumber" });
            return listContent;
        }

        private FileStream GetStreamToWrite()
        {
            return File.Open(this.filename, FileMode.OpenOrCreate, FileAccess.ReadWrite);
        }

        private static NPOIUserModel.ISheet GetSheet(string filename)
        {
            var excelFile = new FileStream(filename, FileMode.Open, FileAccess.Read);
            var workbook = new HSSFWorkbook(excelFile, true);
            var sheet = workbook.GetSheet("hoja1");
            return sheet;
        }

        private static ExcelWriter CreateExcelWriter()
        {
            var excelWriter = WriterFactory.ExcelWriter();
            return excelWriter;
        }
    }
}