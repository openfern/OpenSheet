using System;

namespace OpenSheet.Logging
{
    /// <summary>
    /// Compatibility shim. OpenSheet writes no log files and sends nothing anywhere, so these settings are stored and
    /// otherwise ignored.
    /// </summary>
    [Obsolete("OpenSheet has no logging; this class only exists so migrated code compiles.")]
    public static class Logger
    {
        /// <summary>Where to log. Stored, never used.</summary>
        public static string LogFilePath { get; set; }

        /// <summary>Which sinks to log to. Stored, never used.</summary>
        public static LoggingModes LoggingMode { get; set; }

        /// <summary>Does nothing: OpenSheet never creates log files.</summary>
        public static void ClearLogFiles() { }

        [Flags]
        public enum LoggingModes
        {
            None = 0, DebugOutputWindow = 1, Console = 2, Custom = 4, File = 8, All = 15,
        }
    }
}
