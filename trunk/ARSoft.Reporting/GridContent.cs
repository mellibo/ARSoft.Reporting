namespace ARSoft.Reporting
{
    public class ListContent : ReportContent
    {
        public string DataSourceExpression { get; set; }

        public override void Write(ExcelWriter excelWriter, object datasource)
        {
            throw new System.NotImplementedException();
        }
    }
}