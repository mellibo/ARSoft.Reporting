namespace ARSoft.Reporting.Tests
{
    using System.IO;

    public class MockWriter : IReportWriter
    {
        private int rowCount;

        private string lastWritedText;

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

        public void StartRender(Stream streamToWrite)
        {
            this.StartRender(streamToWrite, null);
        }

        public void StartRender(Stream streamToWrite, string template)
        {
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
        }

        public void CrLf()
        {
            this.rowCount++;
        }
    }
}