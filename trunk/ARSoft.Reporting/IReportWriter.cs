namespace ARSoft.Reporting
{
    using System.IO;

    public interface IReportWriter
    {
        void StartRender(Stream streamToWrite);

        void StartRender(Stream streamToWrite, string template);

        void EndRender();

        void WriteTextElement(int? x, int? y, string text);

        void CrLf();
    }
}