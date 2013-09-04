namespace ARSoft.Reporting
{
    using System;
    using System.Collections.Generic;
    using System.IO;
    using System.Linq;

    using ARSoft.NPOI;

    using global::NPOI.HSSF.UserModel;
    using global::NPOI.SS.UserModel;

    public class ExcelWriter : IReportWriter
    {
        private RenderContext context;

        private Stream streamToWrite;

        private HSSFWorkbook workbook;

        private ISheet sheet;

        private int lastX;

        private int lastY;

        private List<ICellStyle> styles = new List<ICellStyle>();

        public ExcelWriter()
        {
            this.lastX = -1;
            this.lastY = 0;
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

            this.context = new RenderContext();
            
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
            if (!x.HasValue) x = ++lastX;
            if (!y.HasValue) y = lastY;
            lastX = x.Value;
            lastY = y.Value;
            var row = this.sheet.GetRow(y.Value) ?? this.sheet.CreateRow(y.Value);
            var cell = row.GetCell(x.Value) ?? row.CreateCell(x.Value);
            cell.SetCellValue(text);
        }

        /// <summary>
        /// Carriage Return, Line Feed
        /// </summary>
        public void CrLf()
        {
            this.lastX = -1;
            this.lastY++;
        }

        public RenderContext Context
        {
            get
            {
                return this.context;
            }
        }

        public int LastX
        {
            get
            {
                return this.lastX;
            }
        }

        public void SetCurrentX(int x)
        {
            this.lastX = x;
        }

        public void SetCurrentY(int y)
        {
            this.lastY = y;
        }

        public void StartRow(string itemTemplate)
        {
            if (!string.IsNullOrWhiteSpace(itemTemplate))
            {
                var sheetTemplate = workbook.GetSheet("template");

                var rowTemplate = sheetTemplate.GetRow(int.Parse(itemTemplate));
                if (styles.Count == 0)
                {
                    foreach (var cell in rowTemplate.Cells)
                    {
                        var style = sheetTemplate.Workbook.CreateCellStyle();
                        style.CloneStyleFrom(cell.CellStyle);
                        styles.Add(style);
                    }
                }

                this.sheet.CopyRow(sheetTemplate, int.Parse(itemTemplate), this.lastY);

                var rowCopied = this.sheet.GetRow(this.lastY);
                for (int j = 0; j < styles.Count; j++)
                {
                    rowCopied.Cells.First(x => x.ColumnIndex == rowTemplate.Cells[j].ColumnIndex).CellStyle = styles[j];
                }
            }
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