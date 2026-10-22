namespace OpenSheet
{
    /// <summary>
    /// OpenSheet is free software and needs no license key. This type only exists so code written against
    /// libraries that require a key keeps compiling; every member is inert.
    /// </summary>
    public static class License
    {
        public static bool IsLicensed => true;
        public static string LicenseKey { get; set; }
        public static void DisableAppAnalytics() { }
        public static bool IsValidLicense(string licenseKey) => true;
    }
}
