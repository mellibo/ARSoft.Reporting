namespace ARSoft.Reporting.Tests
{
    using NUnit.Framework;

    using SharpTestsEx;

    [TestFixture]
    public class ReportTests
    {
        [Test]
        public void ReportContienteUnaListaDeContenidos()
        {
            // arrange

            // act
            var report = new Report();

            // assert
            report.Contents.Should().Not.Be.Null();
        }

    }

    public class Report
    {
        public object Contents
        {
            get
            {
                return new object();
            }
        }
    }
}