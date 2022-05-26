using System;

namespace KpiISO.Data.Model
{
    public class KpiIsoModel
    {
    }

    public class PrmGetRpt
    {
        public string in_lob { get; set; }
        public string in_projid { get; set; }
        public string in_period { get; set; }
        public string in_projtype { get; set; }
        public string type_rpt { get; set; }
    }

    public class DataKpiHO
    {
        public string io_code { get; set; }
        public string proj_code { get; set; }
        public string proj_name { get; set; }
        public string proj_abbr_code { get; set; }
        public int? date_diff { get; set; }
        public string handover_type { get; set; }
        public int? month { get; set; }
        public int? year { get; set; }
        public DateTime? handover_date { get; set; }
        public DateTime? contract_transfer_date { get; set; }
        public bool cancel { get; set; }
        public Guid lob_id { get; set; }
        public Guid proj_id { get; set; }
        public string proj_type { get; set; }
    }

    public class DataKpiFix
    {
        public string proj_code { get; set; }
        public string proj_name { get; set; }
        public string io_code { get; set; }
        public string nj_code { get; set; }
        public DateTime? due_date { get; set; }
        public DateTime? postpone_due_date { get; set; }
        public DateTime? nj_closed_date { get; set; }
        public int? diff_date { get; set; }
        public int? years { get; set; }
        public int? months { get; set; }
        public Guid nj_id { get; set; }
        public Guid lob_id { get; set; }
        public Guid proj_id { get; set; }
        public string proj_type { get; set; }
    }

    public class ValuePeriod
    {
        public string monthTH { get; set; }
        public string yearTH { get; set; }
        public int iMonth_current { get; set; }
        public int iYear_current { get; set; }
    }

    public class CountKpiHoPeriod
    {
        public int PlanHo { get; set; } //จำนวนแปลงที่ครบกำหนดโอนกรรมสิทธิ์ทั้งสิ้นในเดือน

        public int ActualHoIn3 { get; set; } //2.1 ลูกค้ารับมอบได้ก่อนกำหนดโอนฯ อย่างน้อย 3 วัน (เฉพาะแปลงที่ครบกำหนดโอนฯ ในเดือนนี้) 

        public int ActualHoInCancel { get; set; } //2.2 จำนวนแปลงที่ครบกำหนดโอนฯ เดือนนี้ อยู่ระหว่างการยกเลิก (ได้รับการรับมอบแทนโดยฝ่ายขาย หรือ ผู้มีอำนาจ)

        public int ActualHoTypeSupalai { get; set; } //2.3 แปลงที่เก็บงานแล้วเสร็จ และมีการรับมอบแทนโดยฝ่ายขาย (เฉพาะที่ครบกำหนดโอนเดือนนี้)

        public int SumAct { get; set; } //ผลรวม : จำนวนแปลงที่รับมอบได้ทั้งสิ้นในเดือนนี้ (ข้อ 2.1 + 2.2 + 2.3 )

        public int ActualHoOver { get; set; } //3. จำนวนแปลงที่ส่งมอบได้แต่เกินระยะเวลาที่กำหนด(ทั้งรับมอบโดยลูกค้าและรับมอบแทนโดยฝ่ายขาย)

        public int ActualHoNull { get; set; } //4. ยอดรับมอบคงค้างในเดือน (ข้อ1 - ผลรวมข้อ 2 - ข้อ 3) (ยกไป ข้อ 1 ของ OBJ เรื่องที่ 2 เดือนถัดไป)

        public double Perc { get; set; } //% ที่ส่งมอบได้ตามเป้าหมายของเดือน (ผลรวมข้อ 2 ¸ ข้อ1)
        public int Month { get; set; }
        public int Year { get; set; }
    }

    public class CountKpiHoLate
    {
        public string ProjName { get; set; } //โครงการ
        public string ProjCount { get; set; } //จำนวน
        public string HoCount { get; set; } //จำนวน(ส่งมอบ)
        public string HoProjCode { get; set; } //เลขที่แปลง(ส่งมอบ)
        public string HoNullCount { get; set; } //จำนวน(คงค้าง)
        public string HoNullProjCode { get; set; } //เลขที่แปลง(คงค้าง)
    }

    public class CountKpiHoNullPeriodLowRise
    {
        public int
            PlanHoNullPreMonth
        {
            get;
            set;
        } //1. จำนวนแปลงที่คงค้างการส่งมอบจากเดือนก่อนๆ(ยกมาจากข้อ 4 ของ Obj เรื่องที่ 1 ของเดือนก่อน)

        public int ActHo { get; set; } //2. จำนวนแปลงคงค้าง ที่สามารถส่งมอบได้ในเดือน
        public int ActHoNull { get; set; } //3. จำนวนแปลงคงค้างที่ยังไม่สามารถส่งมอบได้(ข้อ 1 - 2)
        public double Perc { get; set; } //% ที่ส่งมอบบ้านที่คงค้างได้ในเดือน(ข้อ 2 / ข้อ1)

        public int
            MoniterHoNull
        {
            get;
            set;
        } //4. จำนวนแปลงที่ต้องส่งมอบบ้านคงค้างสะสม ณ เดือนนี้(ข้อ 3 เดือนปัจจุบัน + ข้อ 6 ของเดือนก่อน)

        public int
            MoniterHo
        {
            get;
            set;
        } //5. จำนวนแปลงที่คงค้าง จากเดือนก่อนๆ แต่ส่งมอบได้ในเดือนนี้(คงค้างทุกเดือน รับมอบได้จากข้อ 6 เดือนก่อนหน้า)

        public int HoNull { get; set; } //6. คงเหลือจำนวนแปลงคงค้างสะสมที่ไม่สามารถส่งมอบบ้านได้(ข้อ 4 - 5)
        public int Month { get; set; }
        public int Year { get; set; }
    }

    public class CountKpiHoNullLateLowRise
    {
        public string ProjName { get; set; } //โครงการ
        public string ProjCount { get; set; } //จำนวน
        public string HoNullCount { get; set; } //จำนวน(คงค้าง)
        public string HoNullProjCode { get; set; } //เลขที่แปลง(คงค้าง)
    }

    public class CountKpiFixLowRise
    {
        public int
            PlanFixDuedateInMounth
        {
            get;
            set;
        } //1.1 จำนวนใบแจ้งซ่อมที่กำหนดเสร็จในเดือนที่รายงาน(ไม่นับรวมใบแจ้งซ่อมที่ปิดงานและรายงานผลแล้วเมื่อเดือนก่อนๆ)

        public int
            PlanFixCloseInMonth
        {
            get;
            set;
        } //1.2 จำนวนใบแจ้งซ่อมที่มีกำหนดแล้วเสร็จเดือนต่อๆ ไป แต่สามารถปิดงานได้ก่อนในเดือนที่รายงาน

        public int
            PlanFixHo { get; set; } //1.3 จำนวนใบแจ้งซ่อมที่คงค้างการแก้ไขจากเดือนก่อน(ยกมาจาก "ข้อ 4" ของเดือนก่อนๆ)

        public int
            SumPlan
        {
            get;
            set;
        } //จำนวนใบแจ้งซ่อมที่ต้องแก้ไข้แล้วเสร็จทั้งสิ้นภายในเดือนที่วัดผล(ข้อ 1.1 + ข้อ 1.2 + ข้อ 1.3)

        public int
            ActualDuedateInMounth { get; set; } //2.1 ใบแจ้งซ่อมที่มีกำหนดแล้วเสร็จในเดือนที่รายงานที่สามารถปิดได้

        public int
            ActualCloseInMonth
        {
            get;
            set;
        } //2.2 ใบแจ้งซ่อมที่มีกำหนดเสร็จในเดือนต่อๆ ไป แต่สามารถปิดได้ในเดือนที่รายงานผล

        public int ActualFixHo { get; set; } //2.3 จำนวนใบแจ้งซ่อมที่คงค้างการแก้ไขจากเดือนก่อนๆ แต่ปิดได้ในเดือนนี้

        public int
            SumActual { get; set; } //จำนวนใบแจ้งซ่อมที่สามารถแก้ไข้แล้วเสร็จทั้งสิ้นภายในเดือนที่วัดผล(ข้อ 2.1 + 2.2)

        public int
            FixCloseInMonth
        {
            get;
            set;
        } //3. จำนวนใบแจ้งซ่อมที่ปิดงานได้ในเดือนที่รายงาน แต่เกินกำหนดเวลาที่ตกลงกับลูกค้า

        public int
            FixHo
        {
            get;
            set;
        } //4. จำนวนใบแจ้งซ่อมที่ไม่สามารถแก้ไขแล้วเสร็จตามกำหนดเวลาที่ตกลงกับลูกค้า(คงค้างยกไปเดือนถัดไป)

        public double Perc { get; set; } //% ใบแจ้งซ่อมที่แก้ไขแล้วเสร็จภายในกำหนดเวลาที่ตกลงกับลูกค้า(ผลรวม ข้อ 2 ¸ 1)
        public int Month { get; set; }
        public int Year { get; set; }
    }
    
    public class CountKpiHoNullPeriodRegion
    {
        public int
            PlanHoNullPreMonth
        {
            get;
            set;
        } //1.1 จำนวนแปลงที่คงค้างการส่งมอบจากเดือนก่อน(ยกมาจากข้อ 4 ของ Obj ที่ 1 ของเดือนก่อน)

        public int
            PlanHoOverPreMonth
        {
            get;
            set;
        } //1.2 จำนวนแปลงที่คงค้างการส่งมอบจากเดือนก่อนๆ(ยกมาจากข้อ 3 ของ Obj ที่ 2 ของเดือนก่อน)

        public int SumHoNullPreMonth { get; set; } //รวมจำนวนแปลงคงค้าง ที่ต้องส่งมอบทั้งสิ้นในเดือน(1.1 + 1.2)
        public int ActHo { get; set; } //2. จำนวนแปลงที่ส่งมอบได้
        public int ActHoNull { get; set; } //3. จำนวนแปลงคงค้างที่ไม่สามารถส่งมอบได้
        public double Perc { get; set; } //% ที่ส่งมอบแปลงที่คงค้างได้ของเดือน(ข้อ2 / ผลรวมข้อ1)
        public int Month { get; set; }
        public int Year { get; set; }
    }

    public class CountKpiHoNullLateRegion
    {
        public string ProjName { get; set; }//โครงการ
        public string ProjCount { get; set; }//จำนวน
        public string HoNullCount { get; set; }//จำนวน(คงค้าง)
        public string HoNullProjCode { get; set; }//เลขที่แปลง(คงค้าง)
    }

    public class CountKpiFixRegion
    {
        public int PlanFixHo { get; set; } //1.1 จำนวนใบแจ้งซ่อมคงค้างยกมาจากเดือนก่อนๆ
        public int PlanFixDuedateInMounth { get; set; } //1.2 จำนวนใบแจ้งซ่อมที่ต้องดำเนินการแก้ไขให้แล้วเสร็จในเดือน
        public int SumPlan { get; set; } //รวมใบแจ้งซ่อมที่ต้องแล้วเสร็จในเดือนทั้งสิ้น
        public int ActualCloseInMonth { get; set; } //จำนวนใบแจ้งซ่อมที่ปิดงานได้ภายในระยะเวลาที่กำหนด
        public int FixCloseInMonth { get; set; } //3. จำนวนใบแจ้งซ่อมที่ปิดงานได้ แต่เกินกำหนดเวลาที่ตกลงกับลูกค้า
        public int FixHo { get; set; } //4. จำนวนใบแจ้งซ่อมที่คงค้าง / แก้ไขไม่แล้วเสร็จตามกำหนด(ยกไปเดือนถัดไป)
        public double Perc { get; set; } //% ใบแจ้งซ่อมที่ปิดงานทันกำหนดตามที่ตกลงกับลูกค้า(ข้อ 2 ¸ ผลรวมข้อ 1)
        public int Month { get; set; }
        public int Year { get; set; }
    }
}