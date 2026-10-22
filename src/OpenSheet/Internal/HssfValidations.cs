using System.Collections.Generic;
using NPOI.HSSF.Record;
using NPOI.HSSF.Record.Aggregates;
using NPOI.HSSF.UserModel;

namespace OpenSheet.Internal
{
    /// <summary>
    /// NPOI 2.7 never removes or edits data validation records of XLS sheets (RemoveDataValidation silently does nothing).
    /// This reaches the sheet's validity table directly.
    /// </summary>
    internal static class HssfValidations
    {
        sealed class Collector : RecordVisitor
        {
            public readonly List<DVRecord> Found = new List<DVRecord>();
            public void VisitRecord(Record r) { if (r is DVRecord d) Found.Add(d); }
        }

        public static void RemoveAt(HSSFSheet sheet, int index)
        {
            var table = sheet.Sheet.GetOrCreateDataValidityTable();
            var c = new Collector();
            table.VisitContainedRecords(c);
            if (index < 0 || index >= c.Found.Count) return;
            table.RemoveDataValidation(c.Found[index]);
        }
    }
}
