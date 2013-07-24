namespace ARSoft.Reporting
{
    using System.Collections.Generic;

    public class ReportDefinition
    {
        readonly IList<ReportContent> reportContents = new List<ReportContent>();

        public IEnumerable<ReportContent> Contents
        {
            get
            {
                return this.reportContents;
            }
        }

        public void AddContent(ReportContent content)
        {
            this.reportContents.Add(content);
        }
    }
}