namespace ARSoft.Reporting
{
    public class StaticContent : ReportContent
    {
        public string Text { get; set; }

        public override void Write(IReportWriter writer, object datasource)
        {
            writer.WriteTextElement(X, Y, Text);
        }
    }
}