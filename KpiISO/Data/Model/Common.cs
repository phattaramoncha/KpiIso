using System;

namespace KpiISO.Data.Model
{
    public class Project
    {
        public Guid proj_id { get; set; }
        public string proj_name { get; set; }
    }
    public class Line_Of_Business
    {
        public string lb_id { get; set; }
        public string lb_code { get; set; }
        public string lb_name { get; set; }
    }
    public class Line_Bus
    {
        public Guid lob_id { get; set; }
        public string lob_name { get; set; }
    }
}