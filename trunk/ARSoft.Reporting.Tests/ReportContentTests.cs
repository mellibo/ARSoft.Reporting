using NUnit.Framework;

using SharpTestsEx;

namespace ARSoft.Reporting.Tests
{
    [TestFixture]
    public class ReportContentTests
    {
        [Test]
        public void UnContenidoSePuedePosicionarPorUnaEtiquetaEnLaPlantilla()
        {
            // arrange


            // act


            // assert
            Assert.Inconclusive();
        }

        [Test]
        public void HayUnContenidoQueEsUnaExpresionEvaluadaSobreElModelo()
        {
            // arrange
            var datasource = DatasourceFactory.GetDatasourceSimpleObject();
            var writer = new MockWriter();

            // act
            var content = new ExpressionContent();
            content.Expression = "model.Nombre";
            content.Write(writer, datasource);

            // assert
            writer.LastWritedText.Should().Be.EqualTo(datasource.Nombre);
        }



    }
}