using System;

namespace KpiISO.App_Helpers
{
    public class ConvertDateHepler
    {
        public string convertMonthTHDot(string strMonth)
        {
            string month = string.Empty;

            switch (Convert.ToInt32(strMonth))
            {
                case 1:
                    month = "ม.ค.";
                    break;
                case 2:
                    month = "ก.พ.";
                    break;
                case 3:
                    month = "มี.ค.";
                    break;
                case 4:
                    month = "เม.ย.";
                    break;
                case 5:
                    month = "พ.ค.";
                    break;
                case 6:
                    month = "มิ.ย.";
                    break;
                case 7:
                    month = "ก.ค.";
                    break;
                case 8:
                    month = "ส.ค.";
                    break;
                case 9:
                    month = "ก.ย.";
                    break;
                case 10:
                    month = "ต.ค.";
                    break;
                case 11:
                    month = "พ.ย.";
                    break;
                case 12:
                    month = "ธ.ค.";
                    break;
            }

            return month;
        }
        public string convertDateTH(DateTime? date_)
        {
            string dtTH = string.Empty;

            string format_date = "dd/MM/yyyy";
            if (date_.HasValue)
            {
                DateTime dt = Convert.ToDateTime(date_);
                dtTH = dt.ToString(format_date);

            }
            return dtTH;
        }
        public string ConvertYearTH(string strYear)
        {
            string year = string.Empty;

            if (Convert.ToInt32(strYear) < 2543)
            {
                year = (Convert.ToInt32(strYear) + 543).ToString();
            }

            return year;
        }
    }
}