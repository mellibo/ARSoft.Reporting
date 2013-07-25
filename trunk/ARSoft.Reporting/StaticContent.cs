namespace ARSoft.Reporting
{
    public class StaticContent : ReportContent
    {
        public string Text { get; set; }

        public override void Write(ExcelWriter excelWriter, object datasource)
        {
            excelWriter.WriteTextElement(X, Y, Text);
        }
    }
}