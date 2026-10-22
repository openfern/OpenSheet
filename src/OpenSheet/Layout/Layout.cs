using XL = NPOI.SS.UserModel;

namespace OpenSheet.Layout
{
    /// <summary>Print header or footer text, split into left, center and right parts.</summary>
    public interface IHeaderFooter
    {
        string Left { get; set; }
        string Center { get; set; }
        string Right { get; set; }
    }

    public interface IHeader : IHeaderFooter { }
    public interface IFooter : IHeaderFooter { }

    internal sealed class HeaderFooterAdapter : IHeader, IFooter
    {
        readonly XL.IHeaderFooter _h;
        public HeaderFooterAdapter(XL.IHeaderFooter h) { _h = h; }
        public string Left { get => _h.Left; set => _h.Left = value; }
        public string Center { get => _h.Center; set => _h.Center = value; }
        public string Right { get => _h.Right; set => _h.Right = value; }
    }
}
