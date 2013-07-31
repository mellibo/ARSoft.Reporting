namespace ARSoft.Reporting
{
    public class ListContent : ReportContent
    {
        public string DataSourceExpression { get; set; }
        
        private ReportContentContainer contents;

        public ListContent()
        {
            this.contents = new ReportContentContainer();
        }

        public override void Write(ExcelWriter excelWriter, object datasource)
        {

        }
    }
}