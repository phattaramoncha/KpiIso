using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.IO;
using System.Web;
using System.Web.Script.Serialization;
using System.Web.Services;
using KpiISO.Data.Dao;

namespace KpiISO
{
    public partial class ReportKpiISO : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {
            if (!IsPostBack)
            {
            }
        }

        [WebMethod]
        public static string GetProj(string in_lob, string in_proj_type)
        {
            CommonDao comm = new CommonDao();
            var proj = comm.getProj(in_lob, in_proj_type);

            var json = new JavaScriptSerializer().Serialize(proj);

            return json;
        }

        [WebMethod]
        public static string GetLob()
        {
            CommonDao comm = new CommonDao();
            var lob = comm.getlineBus();

            var json = new JavaScriptSerializer().Serialize(lob);

            return json;
        }

        //[WebMethod(EnableSession = true)]
        [WebMethod]
        public static string GenerateFileExcel(string in_lob, string in_projid, string in_period, string in_projtype,
            string type_rpt)
        {
            return "createRpt.aspx?in_lob=" + in_lob + "&in_projid=" + in_projid + "&in_period=" + in_period +
                   "&in_projtype=" + in_projtype + "&type_rpt=" + type_rpt + ""; //staffID=" + GivenStaffID
        }

        public void createFileExcel(string fileName)
        {
            XSSFWorkbook workbook = new XSSFWorkbook();
            ISheet sht_handover = workbook.CreateSheet("ส่งมอบ");
            ISheet sht_handover_null = workbook.CreateSheet("คงค้าง");
            ISheet sht_fix = workbook.CreateSheet("งานซ่อม");

            #region excute to file excel

            using (var exportData = new MemoryStream())
            {
                HttpContext.Current.Response.Clear();
                workbook.Write(exportData);

                HttpContext.Current.Response.ContentType =
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                HttpContext.Current.Response.AddHeader("Content-Disposition",
                    string.Format("attachment;filename={0}", "รายงานวัตถุประสงค์คุณภาพ(รูปแบบแนวราบ) " + ".xlsx"));
                HttpContext.Current.Response.BinaryWrite(exportData.ToArray());

                HttpContext.Current.Response.End();
            }

            #endregion
        }
    }
}