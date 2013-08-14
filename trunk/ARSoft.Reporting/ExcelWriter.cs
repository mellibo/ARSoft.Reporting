namespace ARSoft.Reporting
{
    using System;
    using System.IO;

    using NPOI.HSSF.UserModel;
    using NPOI.SS.UserModel;

    public class ExcelWriter : IReportWriter
    {
        private Stream streamToWrite;

        private HSSFWorkbook workbook;

        private ISheet sheet;

        private int lastX;
        private int lastY;

        public ExcelWriter()
        {
            this.lastX = 0;
            this.lastY = -1;
        }

        public void StartRender(Stream streamToWrite)
        {
            StartRender(streamToWrite, null);
        }

        public void StartRender(Stream streamToWrite, string template)
        {
            if (streamToWrite == null)
            {
                throw new ArgumentNullException("streamToWrite");
            }

            this.CreateWorkbook(template);

            this.sheet = this.workbook.GetSheet("hoja1") ?? this.workbook.CreateSheet("hoja1");
            this.streamToWrite = streamToWrite;            
        }

        public void EndRender()
        {
            this.workbook.Write(this.streamToWrite);
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

        /// <summary>
        /// Carriage Return, Line Feed
        /// </summary>
        public void CrLf()
        {
            this.lastX++;
            this.lastY = -1;
        }

        private void CreateWorkbook(string template)
        {
            if (string.IsNullOrWhiteSpace(template))
            {
                this.workbook = new HSSFWorkbook();
            }
            else
            {
                using (var fs = new FileStream(template, FileMode.Open, FileAccess.Read))
                {
                    this.workbook = new HSSFWorkbook(fs, true);
                }
            }
        }
    }
}