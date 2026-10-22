using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace OpenSheet.Internal
{
    static class PageSetupFix
    {
        /// <summary>
        /// NPOI writes <c>usePrinterDefaults="0"</c> as soon as repeating rows or columns are set. Excel then ignores the
        /// stored paper size and orientation and shows the printer's own. Dropping the flag makes the stored values count.
        /// </summary>
        public static void Apply(IWorkbook book)
        {
            for (var i = 0; i < book.NumberOfSheets; i++)
            {
                var page = (book.GetSheetAt(i) as XSSFSheet)?.GetCTWorksheet().pageSetup;
                if (page != null && !page.usePrinterDefaults) page.usePrinterDefaults = true;
            }
        }
    }
}
