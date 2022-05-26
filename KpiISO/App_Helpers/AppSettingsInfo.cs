namespace KpiISO.App_Helpers
{
    public static class AppSettingsInfo
    {
        public static string LDAPUrl()
        {
            return Settings.GetValue<string>("LDAPUrl");
        }
        public static string AppName()
        {
            return Settings.GetValue<string>("AppName");
        }
        public static string Domain()
        {
            return Settings.GetValue<string>("Domain");
        }
        public static bool IsSkipAuthen()
        {
            return Settings.GetValue<string>("IsSkipAuthen").Equals("Y");
        }
        public static string HeadOfficeCode()
        {
            return Settings.GetValue<string>("HeadOfficeCode");
        }
        public static string smtpHost()
        {
            return Settings.GetValue<string>("smtpHost");
        }
        public static string smtpPort()
        {
            return Settings.GetValue<string>("smtpPort");
        }
        public static string smtpUser()
        {
            return Settings.GetValue<string>("smtpUser");
        }
        public static string smtpPass()
        {
            return Settings.GetValue<string>("smtpPass");
        }
        public static string IsAllowSendMail()
        {
            return Settings.GetValue<string>("IsAllowSendMail");
        }
        public static string QuotataionPath()
        {
            return Settings.GetValue<string>("QuotataionPath");
        }
        public static string DownloadPath()
        {
            return Settings.GetValue<string>("DownloadPath");
        }
        public static string RejectApptCode()
        {
            return Settings.GetValue<string>("RejectApptCode");
        }
        public static string MediaPath()
        {
            return Settings.GetValue<string>("MediaPath");
        }
        public static string MediaSite()
        {
            return Settings.GetValue<string>("MediaSite");
        }
        public static string ExportPath()
        {
            return Settings.GetValue<string>("ExportPath");
        }
        public static string ExportSite()
        {
            return Settings.GetValue<string>("ExportSite");
        }
        public static string AppVersion()
        {
            return Settings.GetValue<string>("AppVersion");
        }
        public static string MediaPathUpload()
        {
            return Settings.GetValue<string>("MediaPathUpload");
        }
    }
}