using System;

namespace KpiISO
{
    public partial class Authentication : System.Web.UI.Page
    {
        private static readonly log4net.ILog log = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        protected void Page_Load(object sender, EventArgs e)
        {
            //log.Info("Authen");
            //log.Info(IsPostBack);

            if (!IsPostBack)
            {
                //Authen from CM
                try
                {
                    #region MyRegion
                  
                    //bool isSkip = Convert.ToBoolean(ConfigurationManager.AppSettings["SkipAuthen"]);
                    //if (isSkip)
                    //{
                    //    AuthHelper.SignIn("1",
                    //                "SKIP",
                    //                "TAMONWAN.RAN",
                    //                "cm",
                    //                "report",
                    //                "1571188993 ");
                    //}

                    //var request = Request;
                    //var token = request.Form["token"];
                    //var id = request.Form["id"];
                    //var code = request.Form["code"];

                    //if (!AuthHelper.IsAuthenticated())
                    //{
                    //    Dictionary<string, object> payloadArr = null;
                    //    if (token == null)
                    //    {
                    //        RedirectToAuthenErrorPage();
                    //    }
                    //    else
                    //    {
                    //        var certFile = System.Web.HttpContext.Current.Server.MapPath(ConfigurationManager.AppSettings["KeyPath"].ToString());
                    //        AsymmetricKeyParameter asymmetricKeyParameter =
                    //            PublicKeyFactory.CreateKey(File.ReadAllBytes(certFile));

                    //        var rsaKeyParameters = (RsaKeyParameters)asymmetricKeyParameter;
                    //        var rsaParams = DotNetUtilities.ToRSAParameters((RsaKeyParameters)asymmetricKeyParameter);
                    //        var rsaCsp = new RSACryptoServiceProvider();
                    //        rsaCsp.ImportParameters(rsaParams);

                    //        string[] jwtParts = token.Split('.');
                    //        if (jwtParts.Length < 3)
                    //        {
                    //            RedirectToAuthenErrorPage();
                    //        }
                    //        else
                    //        {
                    //            var sha256 = SHA256.Create();
                    //            var hash = sha256.ComputeHash(Encoding.UTF8.GetBytes(jwtParts[0] + '.' + jwtParts[1]));

                    //            var rsaDeformatter = new RSAPKCS1SignatureDeformatter(rsaCsp);
                    //            rsaDeformatter.SetHashAlgorithm("SHA256");

                    //            if (!rsaDeformatter.VerifySignature(hash, FromBase64Url(jwtParts[2])))
                    //            {
                    //                RedirectToAuthenErrorPage();
                    //            }
                    //            else
                    //            {
                    //                byte[] data = FromBase64Url(jwtParts[1]);
                    //                //byte[] data = Convert.FromBase64String(jwtParts[1]);
                    //                var payload = Encoding.UTF8.GetString(data);
                    //                payloadArr = new JavaScriptSerializer().Deserialize<Dictionary<string, object>>(payload);
                    //                //Check for time expired claim or other claims
                    //                if (payloadArr.ContainsKey("iss"))
                    //                {
                    //                    if (!payloadArr["iss"].ToString().ToLower().Equals(ConfigurationManager.AppSettings["System"].ToString()))
                    //                    {
                    //                        RedirectToAuthenErrorPage();
                    //                    }
                    //                }
                    //                else
                    //                {
                    //                    RedirectToAuthenErrorPage();
                    //                }
                    //            }
                    //        }

                    //        AuthHelper.SignIn(id,
                    //            code,
                    //            payloadArr["sub"].ToString(),
                    //            payloadArr["iss"].ToString(),
                    //            payloadArr["aud"].ToString(),
                    //            payloadArr["exp"].ToString());
                    //    }
                    //}

                    //CommonDao cmDao = new CommonDao();
                    //var menu = cmDao.GetMenu();
                    //string reDir = menu.Where(w => w.ReportId.Equals(Guid.Parse(id))).Select(s => s.MenuRedirectedUrl).FirstOrDefault();
                    ////log.Info("Redir = " + reDir);
                    //RedirectToPage(reDir);
                    #endregion
                }
                catch (Exception ex)
                {
                    log.Error(ex.Message);
                    RedirectToAuthenErrorPage();
                }
            }
        }

        private static byte[] FromBase64Url(string base64Url)
        {
            string padded = base64Url.Length % 4 == 0
                    ? base64Url : base64Url + "====".Substring(base64Url.Length % 4);
            string base64 = padded.Replace("_", "/")
                                      .Replace("-", "+");
            return Convert.FromBase64String(base64);
        }

        private void RedirectToAuthenErrorPage()
        {
            Response.Redirect("Error/AuthenError.aspx");
        }

        private void RedirectToPage(string Redir)
        {
            Response.Redirect(Redir, false);
        }
    }
}