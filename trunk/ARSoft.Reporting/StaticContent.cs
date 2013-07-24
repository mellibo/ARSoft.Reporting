namespace ARSoft.Reporting
{
    public class StaticContent : ReportContent
    {
        public string Text { get; set; }

        public override string GetText()
        {
            return Text;
        }
    }
}