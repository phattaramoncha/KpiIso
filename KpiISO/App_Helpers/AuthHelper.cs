using KpiISO.Data.Model;
using System;
using System.Collections.Generic;
using System.DirectoryServices;
using System.Web;

namespace KpiISO.App_Helpers
{
    public static class AuthHelper
    {
        public static bool SignIn(string userName, string password)
        {
            try
            {
                string domainUser = userName + "@" + AppSettingsInfo.Domain();
                //DirectoryEntry entry = new DirectoryEntry(AppSettingsInfo.LDAPUrl(), domainUser, password);
                DirectoryEntry entry = new DirectoryEntry(AppSettingsInfo.LDAPUrl(), domainUser, password);
                object obj = entry.NativeObject;
                DirectorySearcher search = new DirectorySearcher(entry);

                search.Filter = "(SAMAccountName=" + userName + ")";
                search.PropertiesToLoad.Add("cn");

                SearchResult result = search.FindOne();
                if (result == null)
                {
                    return false;
                }

                entry.Close();
                entry = null;

                return true;
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
                //return false;
            }
        }
        #region MyRegion

        //public static void SignOut()
        //{
        //    HttpContext.Current.Session["UserInfo"] = null;
        //    HttpContext.Current.Session.Abandon();
        //}
        public static bool IsAuthenticated()
        {
            return GetLoggedInUserInfo() != null;
        }
        //public static bool IsAdmin()
        //{
        //    return AuthHelper.GetLoggedInUserInfo().Where(u => u.IsAdmin.Equals(1)).Any();
        //}
        //public static bool IsManager(string ProjectId)
        //{
        //    if (!String.IsNullOrEmpty(AuthHelper.GetLoggedInUserInfo().FirstOrDefault().ProjectId))
        //        return AuthHelper.GetLoggedInUserInfo().Where(u => !String.IsNullOrEmpty(ProjectId) && ProjectId.Contains(u.ProjectId) && u.IsManager.Equals(1)).Any();
        //    return false;
        //}
        //public static bool IsEmp()
        //{
        //    return AuthHelper.GetLoggedInUserInfo().Where(u => u.IsEmp.Equals(1)).Any();
        //}
        //public static bool IsMarketing()
        //{
        //    return AuthHelper.GetLoggedInUserInfo().Where(u => u.IsManager.Equals(0) && u.IsEmp.Equals(0) && u.IsAdmin.Equals(0)).Any();
        //}
        //public static bool IsAds()
        //{
        //    return AuthHelper.GetLoggedInUserInfo().Where(u => u.IsAds.Equals(1)).Any();
        //}
        public static List<UserInfo> GetLoggedInUserInfo()
        {
            return HttpContext.Current.Session["UserInfo"] as List<UserInfo>;
        }
        //public static List<string> GetUserRoleCode()
        //{
        //    List<string> role = new List<string>();

        //    //if (IsAuthenticated())
        //    //{
        //    //    foreach (var r in GetLoggedInUserInfo())
        //    //    {
        //    //        role.Add(r.PermissionCode.Trim());
        //    //    }
        //    //}

        //    return role;
        //}

        #endregion
    }
}