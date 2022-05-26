using KpiISO.Data.Dao;
using KpiISO.Data.Model;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Linq;

namespace KpiISO.App_Helpers
{
    public class BindingData
    {
        public static ValuePeriod setupPeriod(PrmGetRpt prm)
        {
            #region period 
            string year_current = prm.in_period.Substring(0, 4);
            string month_current = prm.in_period.Substring(5, 1);

            //string period_to = year_current + month_current;
            ConvertDateHepler cvDate = new ConvertDateHepler();
            ValuePeriod valuePeriod = new ValuePeriod();
            valuePeriod.monthTH = cvDate.convertMonthTHDot(month_current);
            valuePeriod.yearTH = cvDate.ConvertYearTH(year_current);
            #endregion

            valuePeriod.iMonth_current = Convert.ToInt32(month_current);
            valuePeriod.iYear_current = Convert.ToInt32(year_current);
 
            return valuePeriod;
        }
        public static void bindDataHo(ref XSSFWorkbook workbook, ref ISheet sheet, PrmGetRpt prm)
        {
            var valuePeriod = setupPeriod(prm); 
 
            ReportDao rptDao = new ReportDao();
            var data = rptDao.GetKpiHO(prm)
                 .Where(x => x.month == valuePeriod.iMonth_current);

            var rowIndex = 0;
            var colIndex = 0;

            IRow row = sheet.CreateRow(rowIndex);

            #region  header    
            ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, colIndex++, "io_code", HorizontalAlignment.Center, VerticalAlignment.Center);
            ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, colIndex++, "project_code", HorizontalAlignment.Center, VerticalAlignment.Center);
            ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, colIndex++, "project_name", HorizontalAlignment.Center, VerticalAlignment.Center);
            ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, colIndex++, "project_abbr_code", HorizontalAlignment.Center, VerticalAlignment.Center);
            ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, colIndex++, "date_diff", HorizontalAlignment.Center, VerticalAlignment.Center);
            ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, colIndex++, "handover_type", HorizontalAlignment.Center, VerticalAlignment.Center);
            ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, colIndex++, "month", HorizontalAlignment.Center, VerticalAlignment.Center);
            ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, colIndex++, "year", HorizontalAlignment.Center, VerticalAlignment.Center);
            ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, colIndex++, "handover_date", HorizontalAlignment.Center, VerticalAlignment.Center);
            ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, colIndex++, "contract_transfer_date", HorizontalAlignment.Center, VerticalAlignment.Center);
            ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, colIndex++, "in_cancel", HorizontalAlignment.Center, VerticalAlignment.Center);

            rowIndex++;

            #endregion
            #region data
            foreach (var list in data)
            {
                colIndex = 0;
                row = sheet.CreateRow(rowIndex++);

                ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, colIndex++, list.io_code, HorizontalAlignment.Left, VerticalAlignment.Center);
                ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, colIndex++, list.proj_code, HorizontalAlignment.Left, VerticalAlignment.Center);
                ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, colIndex++, list.proj_name, HorizontalAlignment.Left, VerticalAlignment.Center);
                ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, colIndex++, list.proj_abbr_code, HorizontalAlignment.Left, VerticalAlignment.Center);
                ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, colIndex++, list.date_diff, HorizontalAlignment.Left, VerticalAlignment.Center);
                ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, colIndex++, list.handover_type, HorizontalAlignment.Left, VerticalAlignment.Center);
                ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, colIndex++, list.month, HorizontalAlignment.Left, VerticalAlignment.Center);
                ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, colIndex++, list.year, HorizontalAlignment.Left, VerticalAlignment.Center);
                ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, colIndex++, list.handover_date, HorizontalAlignment.Left, VerticalAlignment.Center);
                ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, colIndex++, list.contract_transfer_date, HorizontalAlignment.Left, VerticalAlignment.Center);
                ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, colIndex++, list.cancel, HorizontalAlignment.Left, VerticalAlignment.Center);

            }

            #endregion

        }
        public static void bindDataFix(ref XSSFWorkbook workbook, ref ISheet sheet, PrmGetRpt prm)
        {
            var valuePeriod = setupPeriod(prm);

            ReportDao rptDao = new ReportDao();
            var data = rptDao.GetKpiFix(prm);
                //.Where(x => x.months == valuePeriod.iMonth_current && x.years == valuePeriod.iYear_current).OrderBy(x => x.proj_code).ThenBy(x => x.nj_code);

            var rowIndex = 0;
            var colIndex = 0;

            IRow row = sheet.CreateRow(rowIndex);

            #region  header  
            ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, colIndex++, "proj_name", HorizontalAlignment.Center, VerticalAlignment.Center);
            ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, colIndex++, "proj_code", HorizontalAlignment.Center, VerticalAlignment.Center);
            ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, colIndex++, "io_code", HorizontalAlignment.Center, VerticalAlignment.Center);
            ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, colIndex++, "nj_code", HorizontalAlignment.Center, VerticalAlignment.Center);
            ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, colIndex++, "due_date", HorizontalAlignment.Center, VerticalAlignment.Center);
            ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, colIndex++, "postpone_due_date", HorizontalAlignment.Center, VerticalAlignment.Center);
            ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, colIndex++, "nj_closed_date", HorizontalAlignment.Center, VerticalAlignment.Center);
            ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, colIndex++, "date_diff", HorizontalAlignment.Center, VerticalAlignment.Center);
            ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, colIndex++, "month", HorizontalAlignment.Center, VerticalAlignment.Center);
            ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, colIndex++, "year", HorizontalAlignment.Center, VerticalAlignment.Center);

            rowIndex++;
            #endregion
            #region data
            foreach (var item in data)
            {
                colIndex = 0;
                row = sheet.CreateRow(rowIndex++);

                ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, colIndex++, item.proj_name, HorizontalAlignment.Left, VerticalAlignment.Center);
                ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, colIndex++, item.proj_code, HorizontalAlignment.Left, VerticalAlignment.Center);
                ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, colIndex++, item.io_code, HorizontalAlignment.Left, VerticalAlignment.Center);
                ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, colIndex++, item.nj_code, HorizontalAlignment.Left, VerticalAlignment.Center);
                ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, colIndex++, item.due_date, HorizontalAlignment.Left, VerticalAlignment.Center);
                ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, colIndex++, item.postpone_due_date, HorizontalAlignment.Left, VerticalAlignment.Center);
                ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, colIndex++, item.nj_closed_date, HorizontalAlignment.Left, VerticalAlignment.Center);
                ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, colIndex++, item.diff_date, HorizontalAlignment.Left, VerticalAlignment.Center);
                ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, colIndex++, item.months, HorizontalAlignment.Left, VerticalAlignment.Center);
                ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, colIndex++, item.years, HorizontalAlignment.Left, VerticalAlignment.Center);

            }
            #endregion

        }
    }
}