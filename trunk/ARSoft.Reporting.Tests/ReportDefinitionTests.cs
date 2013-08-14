namespace ARSoft.Reporting.Tests
{
    using System.Collections.Generic;
    using System.Linq;

    using NUnit.Framework;

    using SharpTestsEx;

    [TestFixture]
    public class ReportDefinitionTests
    {
        [Test]
        public void ReportContieneUnaListaDeContenidos()
        {
            // arrange

            // act
            var report = ReportFactory.GetReport();

            // assert
            var contents = report.Contents.Contents as object;
            contents.Should().Not.Be.Null();
            contents.Should().Be.InstanceOf<IEnumerable<ReportContent>>();
        }

        [Test]
        public void AlContenidoSeLePuedeAgregarUnTextoEstatico()
        {
            // arrange
            var report = ReportFactory.GetReport();

            // act
            var content = new StaticContent();
            content.Text = "lalalala";
            report.Contents.AddContent(content);

            // assert
            report.Contents.Contents.FirstOrDefault(x => x.Equals(content)).Should().Not.Be.Null();
        }
    }
}