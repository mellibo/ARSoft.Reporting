namespace ARSoft.Reporting
{
    public class ReportDefinition
    {
        private ReportContentContainer contents;

        public ReportDefinition()
        {
            this.contents = new ReportContentContainer();
        }

        public ReportContentContainer Contents
        {
            get
            {
                return this.contents;
            }
        }

        public string Template
        {
            get
            {
                throw new System.NotImplementedException();
            }
            set
            {
                throw new System.NotImplementedException();
            }
        }
    }
}