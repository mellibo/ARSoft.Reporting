namespace ARSoft.Reporting.Tests
{
    using System.IO;

    using NUnit.Framework;

    using SharpTestsEx;

    [TestFixture]
    public class ListContentTests
    {
        [Test]
        public void ListContentTieneUnaExpresionParaObtenerElListadoDesdeElModeloYenElWriteDebeIterarLaLista()
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

        private static ListContent ListContentWith1StaticContent()
        {
            var listContent = new Reporting.ListContent();
            listContent.Content.AddContent(new StaticContent());
            return listContent;
        }

        [Test]
        public void ListContentCreaUnaNuevaFilaPorCadaItem()
        {
            // arrange
            var listContent = new ListContent();
            var writer = WriterFactory.MockWriter();

            // act
            listContent.Write(writer, DatasourceFactory.GetDatasourceList());

            // assert
            writer.RowCount.Should().Be.EqualTo(3);
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
            for (int i = 0; i < writer.TextWrited.Count; i++)
            {
                writer.TextWrited[i].Should().Be.EqualTo((i + 1).ToString());
            }
        }
    }
}