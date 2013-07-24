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
            var sheet = workbook.CreateSheet("hoja1");
            foreach (var content in reportDefinition.Contents)
            {
                var row = sheet.GetRow(content.X.Value - 1) ?? sheet.CreateRow(content.X.Value - 1);
                var cell = row.GetCell(content.Y.Value - 1) ?? row.CreateCell(content.Y.Value - 1);
                cell.SetCellValue(content.GetText());
            }

            workbook.Write(st);
        }
    }
}