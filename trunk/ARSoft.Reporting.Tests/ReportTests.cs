namespace ARSoft.Reporting.Tests
{
    using System.Collections.Generic;
    using System.Linq;

    using NUnit.Framework;

    using SharpTestsEx;

    [TestFixture]
    public class ReportTests
    {
        [Test]
        public void ReportContieneUnaListaDeContenidos()
        {
            // arrange

            // act
            var report = new Report();

            // assert
            var contents = report.Contents;
            contents.Should().Not.Be.Null();
            ((object)contents).Should().Be.InstanceOf<IEnumerable<ReportContent>>();
        }

        [Test]
        public void AlContenidoSeLePuedeAgregarUnElementoParaMostrarUnaExpresionDelModelo()
        {
            // arrange
            var report = new Report();


            // act
            var expresionContent = new ExpressionContent();
            expresionContent.Expression = "model.Nombre";
            expresionContent.Position = 2; 
            report.AddContent(expresionContent);

            // assert
            report.Contents.FirstOrDefault(x => x.Equals(expresionContent)).Should().Not.Be.Null();
        }
    }

    public class ExpressionContent : ReportContent
    {
        public string Expression { get; set; }
        
        public int Position { get; set; }
    }

    public class ReportContent
    {
    }

    public class Report
    {
        readonly IList<ReportContent> reportContents = new List<ReportContent>();

        public IEnumerable<ReportContent> Contents
        {
            get
            {
                return reportContents;
            }
        }

        public void AddContent(ExpressionContent expresionContent)
        {
            reportContents.Add(expresionContent);
        }
    }
}