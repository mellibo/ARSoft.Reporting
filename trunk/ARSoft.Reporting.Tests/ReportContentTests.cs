namespace ARSoft.Reporting.Tests
{
    using System;

    using NUnit.Framework;

    using SharpTestsEx;

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
            var writer = WriterFactory.MockWriter();

            // act
            var content = new ExpressionContent();
            content.Expression = "model.Nombre";
            content.Write(writer, datasource);

            // assert
            writer.LastWritedText.Should().Be.EqualTo(datasource.Nombre);
        }

        [Test]
        public void LaExpresionDeExpresionContentPuedeIncluirVariablesDeContextoComoFecha()
        {
            // arrange
            var datasource = DatasourceFactory.GetDatasourceSimpleObject();
            var writer = WriterFactory.MockWriter();

            // act
            var content = new ExpressionContent();
            content.Expression = @"Context.Date.ToString(""dd/MM/yyyy"")";
            content.Write(writer, datasource);

            // assert
            writer.LastWritedText.Should().Be.EqualTo(DateTime.Today.ToString("dd/MM/yyyy"));
        }
    }
}