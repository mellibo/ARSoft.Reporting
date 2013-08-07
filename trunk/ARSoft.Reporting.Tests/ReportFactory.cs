namespace ARSoft.Reporting.Tests
{
    public class ReportFactory
    {
        public static ReportDefinition GetReport()
        {
            var reportDefinition = new ReportDefinition();
            var staticContent = new StaticContent();
            staticContent.Text = "pepe";
            staticContent.X = 5;
            staticContent.Y = 6;
            reportDefinition.Contents.AddContent(staticContent);

            staticContent = new StaticContent();
            staticContent.Text = "pepe 2";
            staticContent.X = 7;
            staticContent.Y = 3;
            reportDefinition.Contents.AddContent(staticContent);

            var expressionContent = new ExpressionContent();
            expressionContent.Expression = "model.NumeroEntero";
            expressionContent.Y = 2;
            expressionContent.X = 4;
            reportDefinition.Contents.AddContent(expressionContent);
            return reportDefinition;
        }

        public static ReportDefinition GetReportSinCoordenadas()
        {
            var reportDefinition = new ReportDefinition();
            var staticContent = new StaticContent();
            staticContent.Text = "pepe";
            reportDefinition.Contents.AddContent(staticContent);

            staticContent = new StaticContent();
            staticContent.Text = "pepe 2";
            reportDefinition.Contents.AddContent(staticContent);

            staticContent = new StaticContent();
            staticContent.Text = "pepe 3";
            reportDefinition.Contents.AddContent(staticContent);

            return reportDefinition;
        }
 
    }
}