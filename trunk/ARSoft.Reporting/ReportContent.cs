namespace ARSoft.Reporting
{
    public abstract class ReportContent
    {
        public int? X { get; set; }
        
        public int? Y { get; set; }

        public abstract void Write(IReportWriter excelWriter, object datasource);
    }
}