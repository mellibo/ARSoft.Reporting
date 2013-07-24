namespace ARSoft.Reporting
{
    using System.Collections;
    using System.IO;

    public class ExcelRenderer
    {
        public void Render(IEnumerable datasource, ReportDefinition reportDefinition, string filename)
        {
            File.Create(filename);
        }
    }
}