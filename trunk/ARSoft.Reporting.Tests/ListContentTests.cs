namespace ARSoft.Reporting.Tests
{
    using System;
    using System.IO;
    using System.Linq;

    using global::NPOI.HSSF.UserModel;

    using global::NPOI.SS.UserModel;

    using NUnit.Framework;

    using SharpTestsEx;

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
        public void ConListAnidadosDebeMantenerElItemNumberDeCadaListaIndependiente()
        {
            // arrange
            var listContentMaster = new ListContent();
            listContentMaster.Content.AddContent(new ExpressionContent { Expression = "Context.ItemNumber" });
            var listContentChild = new ListContent();
            listContentMaster.Content.AddContent(listContentChild);
            listContentChild.Content.AddContent(new ExpressionContent { Expression = "Context.ItemNumber" });
            var datasourceList = DatasourceFactory.GetDatasourceList();
            listContentChild.DataSource = datasourceList;
            var writer = WriterFactory.MockWriter();

            // act
            listContentMaster.Write(writer, datasourceList);

            // assert
            int textWritedCounter = 0;
            for (int i = 0; i < datasourceList.Count; i++)
            {
                Console.WriteLine(textWritedCounter.ToString() + ": " + writer.WritedElements.ToArray()[textWritedCounter].Text);
                writer.WritedElements.ToArray()[textWritedCounter].Text.Should().Be.EqualTo((i + 1).ToString());
                textWritedCounter++;
                for (int j = 0; j < datasourceList.Count; j++)
                {
                    Console.WriteLine(textWritedCounter.ToString() + ": " + writer.WritedElements.ToArray()[textWritedCounter].Text);
                    writer.WritedElements.ToArray()[textWritedCounter].Text.Should().Be.EqualTo((j + 1).ToString());
                    textWritedCounter++;
                }
            }
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