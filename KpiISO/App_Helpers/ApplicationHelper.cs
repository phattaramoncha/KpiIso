namespace KpiISO.App_Helpers
{
    public class ApplicationUser
    {
        public string Id { get; set; }
        public string Code { get; set; }
        public string Subject { get; set; }
        public string Issuer { get; set; }
        public string Audience { get; set; }
        public string ExpirationTime { get; set; }
    }
    #region MyRegion
    //public static class AuthHelper
    //{
    //    public static bool SignIn(string Id, string Code, string Subject, string Issuer, string Audience, string ExpirationTime)
    //    {
    //        HttpContext.Current.Session["User"] = CreateDefualtUser(Id, Code, Subject, Issuer, Audience, ExpirationTime);
    //        return true;
    //    }
    //    public static void SignOut()
    //    {
    //        HttpContext.Current.Session["User"] = null;
    //    }
    //    public static bool IsAuthenticated()
    //    {
    //        return GetLoggedInUserInfo() != null;
    //    }

    //    public static ApplicationUser GetLoggedInUserInfo()
    //    {
    //        return HttpContext.Current.Session["User"] as ApplicationUser;
    //    }
    //    private static ApplicationUser CreateDefualtUser(string Id, string Code, string Subject, string Issuer, string Audience, string ExpirationTime)
    //    {
    //        return new ApplicationUser
    //        {
    //            Id = Id,
    //            Code = Code,
    //            Subject = Subject,
    //            Issuer = Issuer,
    //            Audience = Audience,
    //            ExpirationTime = ExpirationTime
    //        };
    //    }
    //}
    #endregion
}