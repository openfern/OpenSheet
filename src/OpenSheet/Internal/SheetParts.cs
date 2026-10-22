using System.Linq;
using System.Reflection;
using NPOI;
using NPOI.XSSF.Model;
using NPOI.XSSF.UserModel;
using XL = NPOI.SS.UserModel;

namespace OpenSheet.Internal
{
    /// <summary>
    /// NPOI removes a deleted worksheet's own part but leaves the parts it owned (drawing, charts, pictures, comments, tables) in the
    /// package. They are written out anyway, so the deleted sheet's chart data stays in the file, and the next sheet that gets a drawing is
    /// handed the stale one together with the dead sheet's relationships. Releasing them first, from the bottom up, lets NPOI drop each part
    /// once nothing else refers to it (a picture shared with another sheet stays).
    /// </summary>
    internal static class SheetParts
    {
        // protected internal in NPOI 2.7.6, which this library pins
        static readonly MethodInfo RemoveRelation = typeof(POIXMLDocumentPart).GetMethod("RemoveRelation",
            BindingFlags.Instance | BindingFlags.NonPublic | BindingFlags.Public, null, new[] { typeof(POIXMLDocumentPart), typeof(bool) }, null);

        public static void Release(XL.ISheet sheet)
        {
            if (RemoveRelation == null || !(sheet is XSSFSheet x)) return;
            foreach (var rel in x.RelationParts.ToList())
            {
                var part = rel.DocumentPart;
                if (part is XSSFDrawing || part is XSSFVMLDrawing || part is CommentsTable || part is XSSFTable) Detach(x, part);
            }
        }

        static void Detach(POIXMLDocumentPart parent, POIXMLDocumentPart child)
        {
            foreach (var r in child.RelationParts.ToList()) Detach(child, r.DocumentPart);
            // best effort: a part that cannot be released stays in the package exactly as it did before
            try { RemoveRelation.Invoke(parent, new object[] { child, true }); }
            catch (TargetInvocationException) { }
        }
    }
}
