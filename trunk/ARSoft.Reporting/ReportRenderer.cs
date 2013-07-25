namespace ARSoft.Reporting
{
    public class ReportRenderer
    {
        private readonly ExcelWriter excelWriter;

        public ReportRenderer(ExcelWriter excelWriter)
        {
            this.excelWriter = excelWriter;
        }

        public void Render(object datasource, ReportDefinition reportDefinition, string filename)
        {
            this.excelWriter.StartRender(filename);
            
            foreach (var content in reportDefinition.Contents.Contents)
            {
                content.Write(this.excelWriter, datasource);
            }

            this.excelWriter.EndRender();
        }
    }
}