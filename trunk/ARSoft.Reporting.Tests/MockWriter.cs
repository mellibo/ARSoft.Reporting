namespace ARSoft.Reporting.Tests
{
    using System.Collections.Generic;
    using System.Dynamic;
    using System.IO;

    using System.Linq;

    public class MockWriter : IReportWriter
    {
        // TODO MockWriter tiene logica de manejo de xy que debe estar en un SUT en lugar de un mock
        private readonly RenderContext context;

        private int lastX;

        private int lastY;

        private List<TextElement> textWrited;
        
        private int rowCount;

        public MockWriter(RenderContext context)
        {
            this.context = context;
            this.lastX = -1;
            this.lastY = 0;
        }

        public int WriteCount
        {
            get
            {
                return this.textWrited.Count;
            }
        }

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
                return this.textWrited.Last().Text;
            }
        }

        public IEnumerable<TextElement> WritedElements
        {
            get
            {
                return this.textWrited.ToArray();
            }
        }

        public IEnumerable<string> TextWrited
        {
            get
            {
                return this.textWrited.Select(x => x.Text);
            }
        }

        public void StartRender(Stream streamToWrite)
        {
            this.StartRender(streamToWrite, null);
        }

        public void StartRender(Stream streamToWrite, string template)
        {
            this.textWrited = new List<TextElement>();
            this.rowCount = 0;
        }

        public void EndRender()
        {
            
        }

        public void WriteTextElement(int? x, int? y, string text)
        {
            if (!x.HasValue) x = ++lastX;
            if (!y.HasValue) y = lastY;
            lastX = x.Value;
            lastY = y.Value;

            this.textWrited.Add(new TextElement{ Text = text, X = x, Y = y });
        }

        public void CrLf()
        {
            this.rowCount++;
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
            
        }
    }

    public class TextElement
    {
        public string Text { get; set; }

        public int? X { get; set; }

        public int? Y { get; set; }
    }
}