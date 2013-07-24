namespace ARSoft.Reporting
{
    public class ExpressionContent : ReportContent
    {
        public string Expression { get; set; }
        
        public int Position { get; set; }

        public override string GetText()
        {
            throw new System.NotImplementedException();
        }
    }
}