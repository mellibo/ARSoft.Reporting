namespace ARSoft.Reporting.Tests
{
    using System.IO;

    using NUnit.Framework;

    using SharpTestsEx;

    using System.Linq;

    using global::NPOI.HSSF.UserModel;
    using global::NPOI.SS.UserModel;

    [TestFixture]
    public class ListContentTests
    {
        [Test]
        public void ListContentEnElWriteDebeIterarLaLista()
        {
            // arrange
            var listContent = ListContentWith1StaticContent();
            var writer = WriterFactory.MockWriter();

            // act
            var datasourceList = DatasourceFactory.GetDatasourceList();
            listContent.Write(writer, datasourceList);

            // assert
            writer.WriteCount.Should().Be.EqualTo(datasourceList.Count);
        }

        [Test]
        public void ListContentDatasourceConUnaExpresionParaObtenerElListadoDesdeElModelo()
        {
            // arrange
            var listContent = ListContentWith1StaticContent();
            var writer = WriterFactory.MockWriter();

            // act
            var datasource = DatasourceFactory.GetDatasourceSimpleObject();
            listContent.DataSourceExpression = "model.Hijos";
            listContent.Write(writer, datasource);

            // assert
            writer.WriteCount.Should().Be.EqualTo(datasource.Hijos.Count());
        }

        private static ListContent ListContentWith1StaticContent()
        {
            var listContent = new ListContent();
            listContent.Content.AddContent(new StaticContent());
            return listContent;
        }

        [Test]
        public void ListContentCreaUnaNuevaFilaPorCadaItem()
        {
            // arrange
            var listContent = ListContentWith1StaticContent();
            var writer = WriterFactory.MockWriter();

            // act
            listContent.Write(writer, DatasourceFactory.GetDatasourceList());

            // assert
            writer.RowCount.Should().Be.EqualTo(3);
        }

        [Test]
        public void ListContentEmpiezaARenderizarDesdeUnaPosicionXY()
        {
            // arrange
            var listContent = ListContentWith1StaticContent();
            var writer = WriterFactory.MockWriter();

            // act
            listContent.X = 5;
            listContent.Y = 2;
            listContent.Write(writer, DatasourceFactory.GetDatasourceList());

            // assert
            writer.WritedElements.First().X.Should().Be.EqualTo(listContent.X);
            writer.WritedElements.First().Y.Should().Be.EqualTo(listContent.Y);
        }

        [Test]
        public void ListContentPuedeIterarHorizontalmente()
        {
            // arrange
            var listContent = ListContentWith1StaticContent();
            var writer = WriterFactory.MockWriter();

            // act
            listContent.Direction = DirectionEnum.Horinzontal;
            var datasourceList = DatasourceFactory.GetDatasourceList();
            listContent.Write(writer, datasourceList);

            // assert
            writer.RowCount.Should().Be.EqualTo(0);
            writer.WriteCount.Should().Be.EqualTo(datasourceList.Count);
        }

        [Test]
        public void ListContentPuedeRenderizarElNumeroDeItemAIterar()
        {
            // arrange
            var listContent = new ListContent();
            listContent.Content.AddContent(new ExpressionContent { Expression = "Context.ItemNumber" });
            var writer = WriterFactory.MockWriter();

            // act
            var datasourceList = DatasourceFactory.GetDatasourceList();
            listContent.Write(writer, datasourceList);

            // assert
            for (int i = 0; i < writer.TextWrited.Count(); i++)
            {
                writer.WritedElements.ToArray()[i].Text.Should().Be.EqualTo((i + 1).ToString());
            }
        }

        [Test]
        public void ListContentPuedeDefinirUnItemTemplateParaUnWriter()
        {
            // arrange
            var listContent = ListContentWith1StaticContent();
            listContent.X = 5;
            listContent.Content.AddContent(new ExpressionContent { Expression = "Context.ItemNumber", X  = 2 });
            listContent.Content.AddContent(new ExpressionContent { Expression = "model.Nombre", X = 3 });
            listContent.Content.AddContent(new ExpressionContent { Expression = "model.Fecha", X = 4 });
            var writer = WriterFactory.ExcelWriter();
            var datasourceList = DatasourceFactory.GetDatasourceList();

            // act
            var stream = File.Open("planilla.xls", FileMode.OpenOrCreate, FileAccess.ReadWrite);
            listContent.ItemTemplates.Add(writer.GetType(), "6");
            var template = "template.xlt";
            writer.StartRender(stream, template);
            listContent.Write(writer, datasourceList);
            writer.EndRender();
            stream.Close();

            // assert
            var sheet = GetSheet("planilla.xls");
            sheet.GetRow(0).GetCell(5, MissingCellPolicy.CREATE_NULL_AS_BLANK).StringCellValue.Should().Be.EqualTo(
                "template");
            sheet.GetRow(1).GetCell(5, MissingCellPolicy.CREATE_NULL_AS_BLANK).StringCellValue.Should().Be.EqualTo(
                "template");
            sheet.GetRow(2).GetCell(5, MissingCellPolicy.CREATE_NULL_AS_BLANK).StringCellValue.Should().Be.EqualTo(
                "template");
        }

        private static ISheet GetSheet(string filename)
        {
            var excelFile = new FileStream(filename, FileMode.Open, FileAccess.Read);
            var workbook = new HSSFWorkbook(excelFile, true);
            var sheet = workbook.GetSheet("hoja1");
            return sheet;
        }

    }
}