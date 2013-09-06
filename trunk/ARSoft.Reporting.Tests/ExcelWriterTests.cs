namespace ARSoft.Reporting.Tests
{
    using System;
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

        private FileStream streamToWrite;

        private ExcelWriter writer;

        private string template;

        //TODO usar directametne ExcelWriter y no ReportDedinition

        [SetUp]
        public void SetUp()
        {
            this.filename = "report.xls";
            if (File.Exists(this.filename)) File.Delete(this.filename);
            
            writer = CreateExcelWriter();
            this.streamToWrite = this.GetStreamToWrite();

            template = "template.xlt";
            writer.StartRender(this.streamToWrite, template);
        }
        
        [TearDown]
        public void TearDown()
        {
            if (this.streamToWrite != null) this.streamToWrite.Dispose();
        }

        [Test]
        public void UnReportDefinitionSePuedeRenderizarComoExcel()
        {
            // arrange
            var reportDefinition = ReportFactory.GetReport();
            var datasource = DatasourceFactory.GetDatasourceSimpleObject();
            var renderer = CreateReportRenderer();

            // act
            renderer.Render(datasource, reportDefinition, this.streamToWrite);
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
            var renderer = CreateReportRenderer();

            // act
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
            var renderer = CreateReportRenderer();

            // act
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
            var datasource = DatasourceFactory.GetDatasourceList();

            // act
            listContent.Write(writer, datasource);

            // assert
            this.EndRenderCloseStream();
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

        private void EndRenderCloseStream()
        {
            this.writer.EndRender();
            this.streamToWrite.Close();
        }

        [Test]
        public void SePuedeREnderizarSobreUnaPlantilla()
        {
            // arrange

            // act
            writer.WriteTextElement(null, null, "pepe");

            // assert
            this.EndRenderCloseStream();
            var sheet = GetSheet(this.filename);
            sheet.GetRow(0).GetCell(5, NPOIUserModel.MissingCellPolicy.CREATE_NULL_AS_BLANK).StringCellValue.Should().Be.EqualTo(
                "template");
        }

        [Test]
        public void AlRenderizarDebeQuedarActivaLaHojaRenderizada()
        {
            // arrange

            // act
            writer.WriteTextElement(null, null, "pepe");

            // assert
            this.EndRenderCloseStream();
            var excelFile = this.GetStreamToWrite();
            var xl = new HSSFWorkbook(excelFile, true);
            xl.GetSheetAt(xl.ActiveSheetIndex).SheetName.ToLower().Should().Be("hoja1");
        }

        [Test]
        public void ListContentEmpiezaARenderizarDesdeUnaPosicionXY()
        {
            // arrange
            var listContent = ListContentWithRowNumber();

            // act
            listContent.X = 5;
            listContent.Y = 2;
            listContent.Write(writer, DatasourceFactory.GetDatasourceList());

            // assert
            this.EndRenderCloseStream();
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

            // act
            listContent.ItemTemplates.Add(typeof(ExcelWriter), "15");
            listContent.Write(writer, DatasourceFactory.GetDatasourceList());

            // assert
            this.EndRenderCloseStream();
            var sheet = GetSheet(this.filename);
            for (int i = 0; i < DatasourceFactory.GetDatasourceList().Count; i++)
            {
                // en el template en col 5 empieza la serie 1,2,3,4,5 con alineamineto centrado y border THIN
                sheet.GetRow(i).GetCell(5, NPOIUserModel.MissingCellPolicy.CREATE_NULL_AS_BLANK).NumericCellValue.Should().Be.EqualTo(1);
                sheet.GetRow(i).GetCell(6, NPOIUserModel.MissingCellPolicy.CREATE_NULL_AS_BLANK).NumericCellValue.Should().Be.EqualTo(2);
                sheet.GetRow(i).GetCell(7, NPOIUserModel.MissingCellPolicy.CREATE_NULL_AS_BLANK).NumericCellValue.Should().Be.EqualTo(3);
                sheet.GetRow(i).GetCell(8, NPOIUserModel.MissingCellPolicy.CREATE_NULL_AS_BLANK).NumericCellValue.Should().Be.EqualTo(4);
                sheet.GetRow(i).GetCell(9, NPOIUserModel.MissingCellPolicy.CREATE_NULL_AS_BLANK).NumericCellValue.Should().Be.EqualTo(5);
                for (var j = 0; j < 5; j++)
                {
                    sheet.GetRow(i).GetCell(5 + j, NPOIUserModel.MissingCellPolicy.CREATE_NULL_AS_BLANK).CellStyle.Alignment.
                        Should().Be.EqualTo(NPOIUserModel.HorizontalAlignment.CENTER);
                    sheet.GetRow(i).GetCell(5 + j, NPOIUserModel.MissingCellPolicy.CREATE_NULL_AS_BLANK).CellStyle.
                        BorderBottom.Should().Be.EqualTo(NPOIUserModel.BorderStyle.THIN);
                    sheet.GetRow(i).GetCell(5 + j, NPOIUserModel.MissingCellPolicy.CREATE_NULL_AS_BLANK).CellStyle.BorderTop.
                        Should().Be.EqualTo(NPOIUserModel.BorderStyle.THIN);
                }
            }
        }

        [Test]
        public void ListContentSiElItemTemplateNoTieneNadaGeneraInvalidOperationException()
        {
            // arrange
            var listContent = ListContentWithRowNumber();

            // act
            listContent.ItemTemplates.Add(typeof(ExcelWriter), "100");
            Executing.This(() => listContent.Write(writer, DatasourceFactory.GetDatasourceList())).Should().Throw<InvalidOperationException>();

            this.EndRenderCloseStream();
        }

        [Test]
        public void SiElItemTemplateDeListContentTieneCeldasConMergeSeDebeRespetarElMerge()
        {
            // arrange
            var listContent = ListContentWithRowNumber();

            // act
            listContent.ItemTemplates.Add(typeof(ExcelWriter), "15");
            listContent.Write(writer, DatasourceFactory.GetDatasourceList());

            // assert
            this.EndRenderCloseStream();
            var sheet = GetSheet(this.filename);
            for (int i = 0; i < DatasourceFactory.GetDatasourceList().Count; i++)
            {
                sheet.GetRow(i).GetCell(10, NPOIUserModel.MissingCellPolicy.CREATE_NULL_AS_BLANK).NumericCellValue.Should().Be.EqualTo(6);
                sheet.GetRow(i).GetCell(10).IsMergedCell.Should().Be(true);
                sheet.GetRow(i).GetCell(11).IsMergedCell.Should().Be(true);
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

        private static ReportRenderer CreateReportRenderer()
        {
            var excelWriter = CreateExcelWriter();
            var renderer = new ReportRenderer(excelWriter);
            return renderer;
        }
    }
}