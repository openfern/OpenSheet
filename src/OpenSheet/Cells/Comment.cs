using OpenSheet.Internal;
using XL = NPOI.SS.UserModel;

namespace OpenSheet.Cells
{
    /// <summary>A note attached to a cell.</summary>
    public class Comment
    {
        readonly XL.IComment _c;
        readonly XL.ICreationHelper _helper;

        internal Comment(XL.IComment comment, XL.ICreationHelper helper) { _c = comment; _helper = helper; }

        public string Author { get => _c.Author; set => _c.Author = value; }

        public string Content
        {
            get => _c.String?.String;
            set => _c.String = _helper.CreateRichTextString(value ?? "");
        }

        /// <summary>Comments are hidden (shown on hover only) unless this is set.</summary>
        public bool IsVisible { get => _c.Visible; set => _c.Visible = value; }

        public string Address => A1.Cell(_c.Row, _c.Column);
    }
}
