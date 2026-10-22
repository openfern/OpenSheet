using System;
using System.IO;

namespace OpenSheet.Tests
{
    /// <summary>Gives each test its own scratch folder, removed afterwards.</summary>
    public abstract class TestBase : IDisposable
    {
        readonly string _dir = Path.Combine(Path.GetTempPath(), "opensheet-tests", Guid.NewGuid().ToString("N"));

        protected TestBase() { Directory.CreateDirectory(_dir); }

        protected string Tmp(string name) => Path.Combine(_dir, name);

        public void Dispose()
        {
            try { Directory.Delete(_dir, true); } catch (IOException) { /* a leaked temp dir is harmless */ }
        }
    }
}
