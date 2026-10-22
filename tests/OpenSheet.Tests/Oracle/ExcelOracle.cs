using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using Xunit;

namespace OpenSheet.Tests.Oracle
{
    /// <summary>
    /// Skips a test when Microsoft Excel is not installed. These tests drive a real, hidden Excel through COM and compare
    /// what it sees with what OpenSheet wrote. Excel is the specification OpenSheet follows, so it is the right judge.
    /// </summary>
    public sealed class ExcelFactAttribute : FactAttribute
    {
        public ExcelFactAttribute()
        {
            if (!ExcelOracle.Available) Skip = "Microsoft Excel is not installed (these tests need Windows + Excel).";
        }
    }

    [CollectionDefinition("Excel")]
    public sealed class ExcelCollection : ICollectionFixture<ExcelOracle> { }

    /// <summary>One hidden Excel process shared by all oracle tests (they run one at a time).</summary>
    public sealed class ExcelOracle : IDisposable
    {
        public static bool Available { get; } = Probe();

        static bool Probe()
        {
            try { return RuntimeInformation.IsOSPlatform(OSPlatform.Windows) && Type.GetTypeFromProgID("Excel.Application") != null; }
            catch (Exception) { return false; }
        }

        dynamic _app;
        int _pid;
        bool _separatorsChanged, _origSystem;
        string _origDecimal, _origThousands;

        /// <summary>
        /// Excel answers "I am busy" (RPC_E_SERVERCALL_RETRYLATER or RPC_E_CALL_REJECTED) while it starts up or while a previous
        /// run is still closing. That is not a failure of the code under test, so the call is simply retried for a while.
        /// </summary>
        public static T Retry<T>(Func<T> call)
        {
            // Excel can stay busy for minutes when the default printer is a slow network queue (it asks the driver on open),
            // so the patience is generous; a genuine hang is cut off by the test run's own timeout.
            var give_up = DateTime.UtcNow.AddMinutes(4);
            while (true)
            {
                try { return call(); }
                catch (COMException e) when (((uint)e.HResult == 0x8001010A || (uint)e.HResult == 0x80010001) && DateTime.UtcNow < give_up)
                {
                    System.Threading.Thread.Sleep(500);
                }
            }
        }

        public dynamic App
        {
            get
            {
                if (_app != null) return _app;
                _app = Retry(() => Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application")));
                Retry(() => { var ready = _app.Version; return 0; });
                _app.Visible = false;
                _app.DisplayAlerts = false;
                _app.AskToUpdateLinks = false;
                _app.EnableEvents = false;
                // Excel formats numbers with the Windows separators; pin them so expectations do not depend on the test machine's locale.
                // The user's own settings are put back in Dispose.
                try
                {
                    _origSystem = (bool)_app.UseSystemSeparators; _origDecimal = (string)_app.DecimalSeparator; _origThousands = (string)_app.ThousandsSeparator;
                    _app.UseSystemSeparators = false; _app.DecimalSeparator = "."; _app.ThousandsSeparator = ",";
                    _separatorsChanged = true;
                }
                catch (Exception) { /* keep the machine's separators */ }
                try
                {
                    GetWindowThreadProcessId((IntPtr)(int)_app.Hwnd, out var pid);
                    _pid = (int)pid;
                }
                catch (Exception) { /* the PID is only used to clean up a stuck process */ }
                return _app;
            }
        }

        /// <summary>Opens a workbook; a file Excel would have to repair makes this throw.</summary>
        public ExcelDoc Open(string path, string password = null, bool readOnly = true)
        {
            object missing = Type.Missing;
            dynamic wb = Retry(() => App.Workbooks.Open(path, 0, readOnly, missing, password ?? (object)missing));
            return new ExcelDoc(wb);
        }

        /// <summary>A new empty workbook with one sheet.</summary>
        public ExcelDoc New() => new ExcelDoc(Retry(() => App.Workbooks.Add()));

        public void Dispose()
        {
            if (_app == null) return;
            if (_separatorsChanged)
            {
                try
                {
                    _app.DecimalSeparator = _origDecimal; _app.ThousandsSeparator = _origThousands; _app.UseSystemSeparators = _origSystem;
                }
                catch (Exception) { /* nothing more can be done */ }
            }
            try { _app.Quit(); } catch (Exception) { /* already gone */ }
            try { Marshal.FinalReleaseComObject(_app); } catch (Exception) { /* not a COM object any more */ }
            _app = null;
            GC.Collect();
            GC.WaitForPendingFinalizers();
            if (_pid != 0)
            {
                try
                {
                    var p = Process.GetProcessById(_pid);
                    if (!p.WaitForExit(5000)) p.Kill();
                }
                catch (Exception) { /* exited by itself */ }
            }
        }

        [DllImport("user32.dll")]
        static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);
    }

    public sealed class ExcelDoc : IDisposable
    {
        public ExcelDoc(dynamic workbook) { Book = workbook; }

        public dynamic Book { get; }
        public dynamic Sheet(string name) => Book.Worksheets.Item(name);
        public dynamic Sheet(int oneBased) => Book.Worksheets.Item(oneBased);

        public void Dispose()
        {
            try { Book.Close(false); } catch (Exception) { /* already closed */ }
            try { Marshal.FinalReleaseComObject(Book); } catch (Exception) { /* ignore */ }
        }
    }
}
