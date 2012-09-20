using System;
namespace YUMXPSViewer.Utility.XpsDocument
{
    internal interface IXpsPage
    {
        int CurrentPage { get; }
        event EventHandler DocumentLoaded;
        event EventHandler FixedPageChanged;
        void LoadPage(int pageNum);
        int PageCount { get; }
        void SetStream(System.IO.Stream stream);
    }
}
