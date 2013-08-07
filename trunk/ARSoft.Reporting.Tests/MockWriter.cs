namespace ARSoft.Reporting.Tests
{
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

        public void StartRender(string filename)
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

        public void NewRow()
        {
            this.rowCount++;
        }
    }
}