namespace ARSoft.Reporting
{
    using System.Dynamic;
    using System.IO;

    public interface IReportWriter
    {
        void StartRender(Stream streamToWrite);

        void StartRender(Stream streamToWrite, string template);

        void EndRender();

        void WriteTextElement(int? x, int? y, string text);

        void CrLf();

        RenderContext Context { get; }

        void SetCurrentX(int x);
        
        void SetCurrentY(int y);
    }
}