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
            var listContent = new Reporting.ListContent();
            listContent.Content.AddContent(new StaticContent());
            var writer = new MockWriter();
            writer.StartRender(null);

            // act
            var datasourceList = DatasourceFactory.GetDatasourceList();
            listContent.Write(writer, datasourceList);

            // assert
            writer.WriteCount.Should().Be.EqualTo(datasourceList.Count);
        }

        [Test]
        public void ListContentCreaUnaNuevaFilaPorCadaItem()
        {
            // arrange
            var listContent = new ListContent();
            var writer = new MockWriter();

            // act
            listContent.Write(writer, DatasourceFactory.GetDatasourceList());

            // assert
            writer.RowCount.Should().Be.EqualTo(3);
        }

        [Test]
        public void ListContentPuedeIterarHorizontalmente()
        {
            // arrange
            var listContent = new Reporting.ListContent();
            listContent.Content.AddContent(new StaticContent());
            var writer = new MockWriter();
            writer.StartRender(null);

            // act
            listContent.Direction = DirectionEnum.Horinzontal;
            var datasourceList = DatasourceFactory.GetDatasourceList();
            listContent.Write(writer, datasourceList);

            // assert
            writer.RowCount.Should().Be.EqualTo(0);
            writer.WriteCount.Should().Be.EqualTo(datasourceList.Count);
        }

    }
}