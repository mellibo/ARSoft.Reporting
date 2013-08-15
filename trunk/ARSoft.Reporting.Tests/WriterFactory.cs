namespace ARSoft.Reporting.Tests
{
    public class WriterFactory
    {
        public static MockWriter MockWriter()
        {
            var context = new RenderContext();
            var writer = new MockWriter(context);
            writer.StartRender(null);
            return writer;
        }

        public static ExcelWriter ExcelWriter()
        {
            return new ExcelWriter();
        }
    }
}