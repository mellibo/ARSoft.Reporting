namespace ARSoft.Reporting
{
    using System.IO;

    public class ReportRenderer
    {
        private readonly ExcelWriter excelWriter;

        public ReportRenderer(ExcelWriter excelWriter)
        {
            this.excelWriter = excelWriter;
        }

        public void Render(object datasource, ReportDefinition reportDefinition, Stream streamToWrite)
        {
            Render(datasource, reportDefinition, streamToWrite, null);
        }

        public void Render(object datasource, ReportDefinition reportDefinition, Stream streamToWrite, string template)
        {
            this.excelWriter.StartRender(streamToWrite, template);
            
            foreach (var content in reportDefinition.Contents.Contents)
            {
                content.Write(this.excelWriter, datasource);
            }

            this.excelWriter.EndRender();
        }
    }
}