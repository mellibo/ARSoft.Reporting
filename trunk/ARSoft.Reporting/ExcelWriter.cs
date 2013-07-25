namespace ARSoft.Reporting
{
    using System;
    using System.IO;

    using NPOI.HSSF.UserModel;
    using NPOI.SS.UserModel;

    public class ExcelWriter
    {
        private string filename;

        private HSSFWorkbook workbook;

        private ISheet sheet;

        public void StartRender(string filename)
        {
            if (filename == null)
            {
                throw new ArgumentNullException("filename");
            }

            if (this.workbook == null) this.workbook = new HSSFWorkbook();
            if (this.sheet == null) this.sheet = this.workbook.CreateSheet("hoja1");
            this.filename = filename;            
        }

        public void EndRender()
        {
            using (var st = new FileStream(this.filename,FileMode.OpenOrCreate, FileAccess.ReadWrite))
            {
                this.workbook.Write(st);
                st.Close();
            }
        }

        public void WriteTextElement(int? x, int? y, string text)
        {
            var row = this.sheet.GetRow(x.Value - 1) ?? this.sheet.CreateRow(x.Value - 1);
            var cell = row.GetCell(y.Value - 1) ?? row.CreateCell(y.Value - 1);
            cell.SetCellValue(text);
        }
    }
}