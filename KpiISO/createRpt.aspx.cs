using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.IO;
using System.Linq;
using KpiISO.App_Helpers;
using KpiISO.Data.Dao;
using KpiISO.Data.Model;

namespace KpiISO
{
    public partial class CreateRpt : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {
            if (!IsPostBack)
            {
                var prm_in_lob = Request.QueryString["in_lob"];
                var prm_in_projid = Request.QueryString["in_projid"];
                var prm_in_period = Request.QueryString["in_period"];
                var prm_in_projtype = Request.QueryString["in_projtype"];
                var prm_type_rpt = Request.QueryString["type_rpt"];

                createFlie(prm_in_lob, prm_in_projid, prm_in_period, prm_in_projtype, prm_type_rpt);
            }
        }

        public void createFlie(string prm_in_lob, string prm_in_projid, string prm_in_period, string prm_in_projtype,
            string prm_type_rpt)
        {
            try
            {
                PrmGetRpt prm = new PrmGetRpt();

                #region prm

                prm.in_lob = prm_in_lob;
                prm.in_projid = prm_in_projid;
                prm.in_period = prm_in_period;
                prm.in_projtype = prm_in_projtype;
                prm.type_rpt = prm_type_rpt;

                #endregion

                XSSFWorkbook workbook = new XSSFWorkbook();

                ISheet sht_handover = workbook.CreateSheet("ส่งมอบ");
                ISheet sht_handover_null = workbook.CreateSheet("คงค้าง");
                ISheet sht_fix = workbook.CreateSheet("งานซ่อม");
                ISheet sht_data_ho = workbook.CreateSheet("ข้อมูลแนบส่งมอบ");
                ISheet sht_data_fix = workbook.CreateSheet("ข้อมูลแนบงานซ่อม");

                //แนวราบ
                if (prm_in_lob == "9b70169d-07d2-447d-bc5e-6f16049cb589")
                {
                    KpiLowRise.CreateKpiHoLowRise(ref workbook, ref sht_handover, prm);
                    KpiLowRise.CreateKpiHoNullLowRise(ref workbook, ref sht_handover_null, prm);
                    KpiLowRise.CreateKpiFixLowRise(ref workbook, ref sht_fix, prm);
                }
                //ภูมิภาค
                else
                {
                    KpiRegion.CreateKpiHoRegion(ref workbook, ref sht_handover, prm);
                    KpiRegion.CreateKpiHoNullRegion(ref workbook, ref sht_handover_null, prm);
                    KpiRegion.CreateKpiFix(ref workbook, ref sht_fix, prm);
                }

                BindingData.bindDataHo(ref workbook, ref sht_data_ho, prm);
                BindingData.bindDataFix(ref workbook, ref sht_data_fix, prm);

                string typeRpt = string.Empty;

                #region typeRpt

                if (prm_type_rpt == "region")
                {
                    typeRpt = "รูปแบบภูมิภาค";
                }
                else if (prm_type_rpt == "lowrise")
                {
                    typeRpt = "รูปแบบแนวราบ";
                }
                else if (prm_type_rpt == "all")
                {
                    typeRpt = "รูปแบบภูมิภาค (แนวราบ + อาคารสูง)";
                }

                #endregion

                #region lobName

                CommonDao comm = new CommonDao();
                var lobName = comm.getlineBus().Where(x => x.lob_id == Guid.Parse(prm_in_lob)).Select(x => x.lob_name)
                    .FirstOrDefault();

                #endregion

                exportFileExcel(ref workbook, typeRpt, lobName);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public void exportFileExcel(ref XSSFWorkbook workbook, string typeRpt, string lob)
        {
            try
            {
                #region excute to file excel

                using (var exportData = new MemoryStream())
                {
                    Response.Clear();
                    workbook.Write(exportData);

                    Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                    Response.AddHeader("Content-Disposition",
                        string.Format("attachment;filename={0}",
                            "รายงานวัตถุประสงค์คุณภาพ(" + typeRpt + ") " + "" + lob + ".xlsx"));
                    Response.BinaryWrite(exportData.ToArray());

                    Response.End();
                    //Response.Flush(); // Sends all currently buffered output to the client.
                    //HttpContext.Current.Response.SuppressContent = true;  // Gets or sets a value indicating whether to send HTTP content to the client.
                    //callingPage.ApplicationInstance.CompleteRequest(); // Causes ASP.NET to bypass all events and filtering in the HTTP pipeline**
                }

                #endregion
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
    }
}