namespace ARSoft.Reporting
{
    using System.Collections;
    using System.IO;

    using NPOI.HSSF.Model;
    using NPOI.HSSF.UserModel;

    public class ExcelRenderer
    {
        public void Render(IEnumerable datasource, ReportDefinition reportDefinition, Stream st)
        {
            var workbook = new HSSFWorkbook();
            workbook.CreateSheet("hoja1");
            workbook.Write(st);
        }
    }
}