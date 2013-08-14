namespace ARSoft.Reporting.Tests
{
    using System.Collections.Generic;
    using System.Dynamic;
    using System.IO;

    public class MockWriter : IReportWriter
    {
        private readonly RenderContext context;

        private int rowCount;

        private string lastWritedText;

        private List<string> textWrited;

        public MockWriter(RenderContext context)
        {
            this.context = context;
        }

        public int WriteCount { get; set; }

        public int RowCount
        {
            get
            {
                return this.rowCount;
            }
        }

        public string LastWritedText
        {
            get
            {
                return lastWritedText;
            }
        }

        public List<string> TextWrited
        {
            get
            {
                return textWrited;
            }
        }

        public void StartRender(Stream streamToWrite)
        {
            this.StartRender(streamToWrite, null);
        }

        public void StartRender(Stream streamToWrite, string template)
        {
            this.textWrited = new List<string>();
            this.WriteCount = 0;
            this.rowCount = 0;
        }

        public void EndRender()
        {
            
        }

        public void WriteTextElement(int? x, int? y, string text)
        {
            this.lastWritedText = text;
            this.WriteCount++;
            textWrited.Add(text);
        }

        public void CrLf()
        {
            this.rowCount++;
        }

        public RenderContext Context
        {
            get
            {
                return this.context;
            }
        }
    }
}