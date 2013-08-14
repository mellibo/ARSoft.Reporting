namespace ARSoft.Reporting
{
    using System;

    public class RenderContext 
    {
        public DateTime Date 
        {
            get
            {
                return DateTime.Today;
            }
        }

        public int ItemNumber { get; set; }
    }
}