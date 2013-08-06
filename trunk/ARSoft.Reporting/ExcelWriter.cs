namespace ARSoft.Reporting
{
    using System;
    using System.IO;

    using NPOI.HSSF.UserModel;
    using NPOI.SS.UserModel;

    public interface IReportWriter
    {
        void StartRender(string filename);

        void EndRender();

        void WriteTextElement(int? x, int? y, string text);

        void NewRow();
    }

    public class ExcelWriter : IReportWriter
    {
        private string filename;

        private HSSFWorkbook workbook;

        private ISheet sheet;

        private int lastX;
        private int lastY;

        public ExcelWriter()
        {
            this.lastX = 0;
            this.lastY = -1;
        }

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
            using (var st = new FileStream(this.filename, FileMode.OpenOrCreate, FileAccess.ReadWrite))
            {
                this.workbook.Write(st);
                st.Close();
            }
        }

        public void WriteTextElement(int? x, int? y, string text)
        {
            if (!x.HasValue) x = lastX;
            if (!y.HasValue) y = ++lastY;
            lastX = x.Value;
            lastY = y.Value;
            var row = this.sheet.GetRow(x.Value) ?? this.sheet.CreateRow(x.Value);
            var cell = row.GetCell(y.Value) ?? row.CreateCell(y.Value);
            cell.SetCellValue(text);
        }

        public void NewRow()
        {
            this.lastX++;
            this.lastY = -1;
        }
    }
}