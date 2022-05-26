using KpiISO.Data.Dao;
using KpiISO.Data.Model;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;

namespace KpiISO.App_Helpers
{
    public class KpiLowRise
    {
        static string symbol = "❒";

        public static void CreateKpiHoLowRise(ref XSSFWorkbook workbook, ref ISheet sheet, PrmGetRpt prm)
        {
            var valuePeriod = BindingData.setupPeriod(prm);
            var count_ho = GetCountKpiHoPeriod(prm);
            var count_ho_late = GetCountKpiHoLate(prm);

            ExcelHelper.SetupSheet(ref sheet);

            var rowIndex = 0;
            var colIndex = 0;

            IRow row = sheet.CreateRow(rowIndex);
            
            #region add logo
            //String image1 = "assets\\img\\logo-spl.JPG";//Image location address//./images/logo1.png
            string path_img = HttpContext.Current.Server.MapPath("assets\\img\\logo-spl.JPG");//add picture data to this workbook.
            byte[] bytes = File.ReadAllBytes(path_img);
            int pictureIdx = workbook.AddPicture(bytes, PictureType.PNG);

            ICreationHelper helper = workbook.GetCreationHelper();

            // Create the drawing patriarch.  This is the top level container for all shapes.
            IDrawing drawing = sheet.CreateDrawingPatriarch();

            // add a picture shape
            IClientAnchor anchor = helper.CreateClientAnchor();

            //set top-left corner of the picture,
            //subsequent call of Picture#resize() will operate relative to it
            anchor.Col1 = 1;
            anchor.Row1 = 2;

            IPicture pict = drawing.CreatePicture(anchor, pictureIdx);
            pict.Resize(2.5, 2);
            //auto-size picture relative to its top-left corner
            //pict.Resize();
            #endregion

            row = sheet.CreateRow(rowIndex);
            rowIndex++;
            colIndex = 0;

            #region แบบรายงานวัตถุประสงค์คุณภาพ (รูปแบบแนวราบ)

            row = sheet.CreateRow(rowIndex);
            ExcelHelper.createHeader(ref workbook, ref sheet, ref row, colIndex++,
                "แบบรายงานวัตถุประสงค์คุณภาพ (รูปแบบแนวราบ)", HorizontalAlignment.Center, VerticalAlignment.Center);
            sheet.AddMergedRegion(new CellRangeAddress(rowIndex, rowIndex, 0, 21));
            rowIndex++;
            colIndex = 0;

            #endregion

            #region ฝ่ายก่อสร้างแนวราบ

            row = sheet.CreateRow(rowIndex);
            ExcelHelper.createHeader(ref workbook, ref sheet, ref row, colIndex++, "ฝ่ายก่อสร้างแนวราบ",
                HorizontalAlignment.Center, VerticalAlignment.Center);
            sheet.AddMergedRegion(new CellRangeAddress(rowIndex, rowIndex, 0, 21));
            rowIndex++;
            colIndex = 0;

            #endregion

            #region ข้อมูลสิ้นสุด ณ เดือน

            row = sheet.CreateRow(rowIndex);
            ExcelHelper.createHeader(ref workbook, ref sheet, ref row, colIndex++,
                "ข้อมูลสิ้นสุด ณ เดือน " + valuePeriod.monthTH + " " + valuePeriod.yearTH + "",
                HorizontalAlignment.Center, VerticalAlignment.Center);
            sheet.AddMergedRegion(new CellRangeAddress(rowIndex, rowIndex, 0, 21));
            rowIndex++;
            colIndex = 0;

            #endregion

            #region วัตถุประสงค์ด้านคุณภาพที่ : 1   เรื่อง : ต้องส่งมอบบ้านให้ลูกค้าก่อนกำหนดโอนกรรมสิทธิ์ของแต่ละแปลง อย่างน้อย 3 วัน

            row = sheet.CreateRow(rowIndex);
            ExcelHelper.createHeaderDetail(ref workbook, ref sheet, ref row, colIndex++,
                "วัตถุประสงค์ด้านคุณภาพที่ : 1   เรื่อง : ต้องส่งมอบบ้านให้ลูกค้าก่อนกำหนดโอนกรรมสิทธิ์ของแต่ละแปลง อย่างน้อย 3 วัน",
                HorizontalAlignment.Left, VerticalAlignment.Center);
            sheet.AddMergedRegion(new CellRangeAddress(rowIndex, rowIndex, 0, 21));
            rowIndex++;
            colIndex = 0;

            #endregion

            #region เกณฑ์การวัดผล :  ต้องได้  ≥ 99% ในแต่ละเดือน

            row = sheet.CreateRow(rowIndex);
            ExcelHelper.createHeaderDetail(ref workbook, ref sheet, ref row, colIndex++,
                "เกณฑ์การวัดผล :  ต้องได้  ≥ 99% ในแต่ละเดือน", HorizontalAlignment.Left, VerticalAlignment.Center);
            sheet.AddMergedRegion(new CellRangeAddress(rowIndex, rowIndex, 0, 21));
            rowIndex++;
            colIndex = 0;

            #endregion

            #region รายการ

            row = sheet.CreateRow(rowIndex);
            ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, colIndex++, "", HorizontalAlignment.Center,
                VerticalAlignment.Center);
            ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, colIndex++, "รายการ",
                HorizontalAlignment.Center, VerticalAlignment.Center);
            for (int i = 2; i < 9; i++)
            {
                ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, i, "", HorizontalAlignment.Center,
                    VerticalAlignment.Center);
            }

            sheet.AddMergedRegion(new CellRangeAddress(rowIndex, rowIndex, 1, 8));
            colIndex = 9;
            ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, colIndex++, "ม.ค.",
                HorizontalAlignment.Center, VerticalAlignment.Center);
            ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, colIndex++, "ก.พ.",
                HorizontalAlignment.Center, VerticalAlignment.Center);
            ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, colIndex++, "มี.ค.",
                HorizontalAlignment.Center, VerticalAlignment.Center);
            ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, colIndex++, "เม.ย.",
                HorizontalAlignment.Center, VerticalAlignment.Center);
            ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, colIndex++, "พ.ค.",
                HorizontalAlignment.Center, VerticalAlignment.Center);
            ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, colIndex++, "มิ.ย.",
                HorizontalAlignment.Center, VerticalAlignment.Center);
            ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, colIndex++, "ก.ค.",
                HorizontalAlignment.Center, VerticalAlignment.Center);
            ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, colIndex++, "ส.ค.",
                HorizontalAlignment.Center, VerticalAlignment.Center);
            ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, colIndex++, "ก.ย.",
                HorizontalAlignment.Center, VerticalAlignment.Center);
            ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, colIndex++, "ต.ค.",
                HorizontalAlignment.Center, VerticalAlignment.Center);
            ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, colIndex++, "พ.ย.",
                HorizontalAlignment.Center, VerticalAlignment.Center);
            ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, colIndex++, "ธ.ค.",
                HorizontalAlignment.Center, VerticalAlignment.Center);
            ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, colIndex++, "รวม", HorizontalAlignment.Center,
                VerticalAlignment.Center);
            rowIndex++;
            colIndex = 0;

            #endregion

            #region 1.Plan จำนวนแปลงที่ครบกำหนดโอนกรรมสิทธิ์ทั้งสิ้นในเดือน

            row = sheet.CreateRow(rowIndex);
            ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, colIndex++, "1.Plan",
                HorizontalAlignment.Center, VerticalAlignment.Center);
            ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, colIndex++,
                "จำนวนแปลงที่ครบกำหนดโอนกรรมสิทธิ์ทั้งสิ้นในเดือน", HorizontalAlignment.Left, VerticalAlignment.Center);
            for (int i = 2; i < 22; i++)
            {
                ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, i, "", HorizontalAlignment.Center,
                    VerticalAlignment.Center);
            }

            sheet.AddMergedRegion(new CellRangeAddress(rowIndex, rowIndex, 1, 8));

            int sum1 = 0;
            for (int i = 1; i < (valuePeriod.iMonth_current + 1); i++)
            {
                foreach (var item in count_ho)
                {
                    if (item.Month == i)
                    {
                        ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, (i + 8), item.PlanHo,
                            HorizontalAlignment.Center, VerticalAlignment.Center);
                        sum1 += item.PlanHo;
                    }
                }
            }

            ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, 21, sum1, HorizontalAlignment.Center,
                VerticalAlignment.Center);

            rowIndex++;
            colIndex = 0;

            #endregion

            #region 2.Actual 2.1 ลูกค้ารับมอบได้ก่อนกำหนดโอนฯ อย่างน้อย 3 วัน (เฉพาะแปลงที่ครบกำหนดโอนฯ ในเดือนนี้)

            row = sheet.CreateRow(rowIndex);
            ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, colIndex++, "2.Actual",
                HorizontalAlignment.Center, VerticalAlignment.Center);
            ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, colIndex++,
                "2.1 ลูกค้ารับมอบได้ก่อนกำหนดโอนฯ อย่างน้อย 3 วัน (เฉพาะแปลงที่ครบกำหนดโอนฯ ในเดือนนี้)",
                HorizontalAlignment.Left, VerticalAlignment.Center);
            for (int i = 2; i < 22; i++)
            {
                ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, i, "", HorizontalAlignment.Center,
                    VerticalAlignment.Center);
            }

            sheet.AddMergedRegion(new CellRangeAddress(rowIndex, rowIndex, 1, 8));
            int sum2 = 0;
            for (int i = 1; i < (valuePeriod.iMonth_current + 1); i++)
            {
                foreach (var item in count_ho)
                {
                    if (item.Month == i)
                    {
                        ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, (i + 8), item.ActualHoIn3,
                            HorizontalAlignment.Center, VerticalAlignment.Center);
                        sum2 += item.ActualHoIn3;
                    }
                }
            }

            ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, 21, sum2, HorizontalAlignment.Center,
                VerticalAlignment.Center);

            rowIndex++;
            colIndex = 0;

            #endregion

            #region 2.2 จำนวนแปลงที่ครบกำหนดโอนฯ เดือนนี้ อยู่ระหว่างการยกเลิก(ได้รับการรับมอบแทนโดยฝ่ายขาย หรือ ผู้มีอำนาจ)

            row = sheet.CreateRow(rowIndex);
            ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, colIndex++, "", HorizontalAlignment.Center,
                VerticalAlignment.Center);
            ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, colIndex++,
                "2.2 จำนวนแปลงที่ครบกำหนดโอนฯ เดือนนี้ อยู่ระหว่างการยกเลิก(ได้รับการรับมอบแทนโดยฝ่ายขาย หรือ ผู้มีอำนาจ)",
                HorizontalAlignment.Left, VerticalAlignment.Center);
            for (int i = 2; i < 22; i++)
            {
                ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, i, "", HorizontalAlignment.Center,
                    VerticalAlignment.Center);
            }

            sheet.AddMergedRegion(new CellRangeAddress(rowIndex, rowIndex, 1, 8));

            int sum3 = 0;
            for (int i = 1; i < (valuePeriod.iMonth_current + 1); i++)
            {
                foreach (var item in count_ho)
                {
                    if (item.Month == i)
                    {
                        ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, (i + 8), item.ActualHoInCancel,
                            HorizontalAlignment.Center, VerticalAlignment.Center);
                        sum3 += item.ActualHoInCancel;
                    }
                }
            }

            ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, 21, sum3, HorizontalAlignment.Center,
                VerticalAlignment.Center);

            rowIndex++;
            colIndex = 0;

            #endregion

            #region 2.3 แปลงที่เก็บงานแล้วเสร็จ และมีการรับมอบแทนโดยฝ่ายขาย(เฉพาะที่ครบกำหนดโอนเดือนนี้)

            row = sheet.CreateRow(rowIndex);
            ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, colIndex++, "", HorizontalAlignment.Center,
                VerticalAlignment.Center);
            ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, colIndex++,
                "2.3 แปลงที่เก็บงานแล้วเสร็จ และมีการรับมอบแทนโดยฝ่ายขาย(เฉพาะที่ครบกำหนดโอนเดือนนี้)",
                HorizontalAlignment.Left, VerticalAlignment.Center);
            for (int i = 2; i < 22; i++)
            {
                ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, i, "", HorizontalAlignment.Center,
                    VerticalAlignment.Center);
            }

            sheet.AddMergedRegion(new CellRangeAddress(rowIndex, rowIndex, 1, 8));

            int sum4 = 0;
            for (int i = 1; i < (valuePeriod.iMonth_current + 1); i++)
            {
                foreach (var item in count_ho)
                {
                    if (item.Month == i)
                    {
                        ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, (i + 8), item.ActualHoTypeSupalai,
                            HorizontalAlignment.Center, VerticalAlignment.Center);
                        sum4 += item.ActualHoTypeSupalai;
                    }
                }
            }

            ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, 21, sum4, HorizontalAlignment.Center,
                VerticalAlignment.Center);

            rowIndex++;
            colIndex = 0;

            #endregion

            #region ผลรวม : จำนวนแปลงที่รับมอบได้ทั้งสิ้นในเดือนนี้ (ข้อ 2.1 + 2.2 + 2.3 )

            row = sheet.CreateRow(rowIndex);
            ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, colIndex++, "", HorizontalAlignment.Center,
                VerticalAlignment.Center);
            ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, colIndex++,
                "ผลรวม : จำนวนแปลงที่รับมอบได้ทั้งสิ้นในเดือนนี้ (ข้อ 2.1 + 2.2 + 2.3 )", HorizontalAlignment.Left,
                VerticalAlignment.Center);
            for (int i = 2; i < 22; i++)
            {
                ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, i, "", HorizontalAlignment.Center,
                    VerticalAlignment.Center);
            }

            sheet.AddMergedRegion(new CellRangeAddress(rowIndex, rowIndex, 1, 8));

            int sum5 = 0;
            for (int i = 1; i < (valuePeriod.iMonth_current + 1); i++)
            {
                foreach (var item in count_ho)
                {
                    if (item.Month == i)
                    {
                        ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, (i + 8), item.SumAct,
                            HorizontalAlignment.Center, VerticalAlignment.Center);
                        sum5 += item.SumAct;
                    }
                }
            }

            ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, 21, sum5, HorizontalAlignment.Center,
                VerticalAlignment.Center);

            rowIndex++;
            colIndex = 0;

            #endregion

            #region 3. จำนวนแปลงที่ส่งมอบได้แต่เกินระยะเวลาที่กำหนด(ทั้งรับมอบโดยลูกค้าและรับมอบแทนโดยฝ่ายขาย)

            row = sheet.CreateRow(rowIndex);
            ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, colIndex++, "", HorizontalAlignment.Center,
                VerticalAlignment.Center);
            ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, colIndex++,
                "3. จำนวนแปลงที่ส่งมอบได้แต่เกินระยะเวลาที่กำหนด(ทั้งรับมอบโดยลูกค้าและรับมอบแทนโดยฝ่ายขาย)",
                HorizontalAlignment.Left, VerticalAlignment.Center);
            for (int i = 2; i < 22; i++)
            {
                ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, i, "", HorizontalAlignment.Center,
                    VerticalAlignment.Center);
            }

            sheet.AddMergedRegion(new CellRangeAddress(rowIndex, rowIndex, 1, 8));

            int sum6 = 0;
            for (int i = 1; i < (valuePeriod.iMonth_current + 1); i++)
            {
                foreach (var item in count_ho)
                {
                    if (item.Month == i)
                    {
                        ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, (i + 8), item.ActualHoOver,
                            HorizontalAlignment.Center, VerticalAlignment.Center);
                        sum6 += item.ActualHoOver;
                    }
                }
            }

            ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, 21, sum6, HorizontalAlignment.Center,
                VerticalAlignment.Center);

            rowIndex++;
            colIndex = 0;

            #endregion

            #region 4. ยอดรับมอบคงค้างในเดือน (ข้อ1 - ผลรวมข้อ 2 - ข้อ 3)(ยกไป ข้อ 1 ของ OBJ เรื่องที่ 2 เดือนถัดไป)

            row = sheet.CreateRow(rowIndex);
            ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, colIndex++, "", HorizontalAlignment.Center,
                VerticalAlignment.Center);
            ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, colIndex++,
                "4. ยอดรับมอบคงค้างในเดือน (ข้อ1 - ผลรวมข้อ 2 - ข้อ 3)(ยกไป ข้อ 1 ของ OBJ เรื่องที่ 2 เดือนถัดไป)",
                HorizontalAlignment.Left, VerticalAlignment.Center);
            for (int i = 2; i < 22; i++)
            {
                ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, i, "", HorizontalAlignment.Center,
                    VerticalAlignment.Center);
            }

            sheet.AddMergedRegion(new CellRangeAddress(rowIndex, rowIndex, 1, 8));

            int sum7 = 0;
            for (int i = 1; i < (valuePeriod.iMonth_current + 1); i++)
            {
                foreach (var item in count_ho)
                {
                    if (item.Month == i)
                    {
                        ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, (i + 8), item.ActualHoNull,
                            HorizontalAlignment.Center, VerticalAlignment.Center);
                        sum7 += item.ActualHoNull;
                    }
                }
            }

            ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, 21, sum7, HorizontalAlignment.Center,
                VerticalAlignment.Center);

            rowIndex++;
            colIndex = 0;

            #endregion

            #region % ที่ส่งมอบได้ตามเป้าหมายของเดือน (ผลรวมข้อ 2 / ข้อ1)

            row = sheet.CreateRow(rowIndex);
            ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, colIndex++, "%", HorizontalAlignment.Center,
                VerticalAlignment.Center);
            ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, colIndex++,
                "% ที่ส่งมอบได้ตามเป้าหมายของเดือน (ผลรวมข้อ 2 / ข้อ1)", HorizontalAlignment.Left,
                VerticalAlignment.Center);
            for (int i = 2; i < 22; i++)
            {
                ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, i, "", HorizontalAlignment.Center,
                    VerticalAlignment.Center);
            }

            sheet.AddMergedRegion(new CellRangeAddress(rowIndex, rowIndex, 1, 8));

            double sum8 = 0;
            for (int i = 1; i < (valuePeriod.iMonth_current + 1); i++)
            {
                foreach (var item in count_ho)
                {
                    if (item.Month == i)
                    {
                        ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, (i + 8), item.Perc + "%",
                            HorizontalAlignment.Center, VerticalAlignment.Center);
                        sum8 += item.Perc;
                    }
                }
            }

            sum8 = (sum8 / valuePeriod.iMonth_current);
            sum8 = (float) Math.Round((double) sum8);

            ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, 21, sum8 + "%", HorizontalAlignment.Center,
                VerticalAlignment.Center);

            rowIndex++;
            colIndex = 0;

            #endregion

            row = sheet.CreateRow(rowIndex);
            rowIndex++;
            colIndex = 0;

            #region หมายเหตุ : แปลงที่ส่งมอบไม่ทันกำหนด

            row = sheet.CreateRow(rowIndex);
            ExcelHelper.CreateDetail(ref workbook, ref sheet, ref row, colIndex++,
                "หมายเหตุ : แปลงที่ส่งมอบไม่ทันกำหนด", true);
            sheet.AddMergedRegion(new CellRangeAddress(rowIndex, rowIndex, 0, 21));
            rowIndex++;
            colIndex = 0;

            #endregion

            #region แปลงที่ส่งมอบได้แต่ช้ากว่ากำหนด / แปลงที่คงค้างการส่งมอบ

            row = sheet.CreateRow(rowIndex);
            colIndex = 4;
            ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, colIndex++,
                "แปลงที่ส่งมอบได้แต่ช้ากว่ากำหนด / แปลงที่คงค้างการส่งมอบ", HorizontalAlignment.Center,
                VerticalAlignment.Center);
            for (int i = 5; i < 20; i++)
            {
                ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, i, "", HorizontalAlignment.Center,
                    VerticalAlignment.Center);
            }

            sheet.AddMergedRegion(new CellRangeAddress(rowIndex, rowIndex, 4, 19));
            rowIndex++;
            colIndex = 0;

            #endregion

            #region โครงการ จำนวน ปัจจุบันส่งมอบแล้ว ค้างส่งมอบ

            row = sheet.CreateRow(rowIndex);
            colIndex = 4;
            ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, colIndex++, "โครงการ",
                HorizontalAlignment.Center, VerticalAlignment.Center);
            for (int i = 5; i < 12; i++)
            {
                ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, i, "", HorizontalAlignment.Center,
                    VerticalAlignment.Center);
            }

            //โครงการ merged รวม   
            //sheet.AddMergedRegion(new CellRangeAddress(rowIndex, rowIndex, 4, 11));
            colIndex = 12;
            ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, colIndex++, "จำนวน",
                HorizontalAlignment.Center, VerticalAlignment.Center);
            ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, colIndex++, "", HorizontalAlignment.Center,
                VerticalAlignment.Center);
            //จำนวน merged รวม
            //sheet.AddMergedRegion(new CellRangeAddress(rowIndex, rowIndex, 12, 13));
            ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, colIndex++, "ปัจจุบันส่งมอบแล้ว",
                HorizontalAlignment.Center, VerticalAlignment.Center);
            ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, colIndex++, "", HorizontalAlignment.Center,
                VerticalAlignment.Center);
            ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, colIndex++, "", HorizontalAlignment.Center,
                VerticalAlignment.Center);
            sheet.AddMergedRegion(new CellRangeAddress(rowIndex, rowIndex, 14, 16));
            ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, colIndex++, "ค้างส่งมอบ",
                HorizontalAlignment.Center, VerticalAlignment.Center);
            ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, colIndex++, "", HorizontalAlignment.Center,
                VerticalAlignment.Center);
            ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, colIndex++, "", HorizontalAlignment.Center,
                VerticalAlignment.Center);
            sheet.AddMergedRegion(new CellRangeAddress(rowIndex, rowIndex, 17, 19));
            rowIndex++;
            colIndex = 0;

            #endregion

            #region จำนวน เลขที่แปลง จำนวน เลขที่แปลง

            row = sheet.CreateRow(rowIndex);
            for (int i = 4; i < 12; i++)
            {
                ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, i, "", HorizontalAlignment.Center,
                    VerticalAlignment.Center);
            }

            //โครงการ merged รวม   
            sheet.AddMergedRegion(new CellRangeAddress(rowIndex - 1, rowIndex, 4, 11));
            colIndex = 12;
            ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, colIndex++, "", HorizontalAlignment.Center,
                VerticalAlignment.Center);
            ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, colIndex++, "", HorizontalAlignment.Center,
                VerticalAlignment.Center);
            //จำนวน merged รวม
            sheet.AddMergedRegion(new CellRangeAddress(rowIndex - 1, rowIndex, 12, 13));
            ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, colIndex++, "จำนวน",
                HorizontalAlignment.Center, VerticalAlignment.Center);
            ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, colIndex++, "เลขที่แปลง",
                HorizontalAlignment.Center, VerticalAlignment.Center);
            ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, colIndex++, "", HorizontalAlignment.Center,
                VerticalAlignment.Center);
            sheet.AddMergedRegion(new CellRangeAddress(rowIndex, rowIndex, 15, 16));
            ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, colIndex++, "จำนวน",
                HorizontalAlignment.Center, VerticalAlignment.Center);
            ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, colIndex++, "เลขที่แปลง",
                HorizontalAlignment.Center, VerticalAlignment.Center);
            ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, colIndex++, "", HorizontalAlignment.Center,
                VerticalAlignment.Center);
            sheet.AddMergedRegion(new CellRangeAddress(rowIndex, rowIndex, 18, 19));
            rowIndex++;
            colIndex = 0;

            #endregion

            int iSum = 1;
            int sProj = 0;
            int sHO = 0;
            int sHO_NULL = 0;

            if (count_ho_late.Count > 0)
            {
                foreach (var iRowSUM in count_ho_late)
                {
                    #region list detail

                    row = sheet.CreateRow(rowIndex);
                    //rowIndex++; colIndex = 0;

                    #region cell data

                    for (int i = 0; i < 4; i++)
                    {
                        ExcelHelper.CreateDetail(ref workbook, ref sheet, ref row, i, "", true);
                    }

                    //sheet.AddMergedRegion(new CellRangeAddress(rowIndex, rowIndex, 0, 3)); 
                    colIndex = 4;
                    ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, colIndex++,
                        iSum + ". " + iRowSUM.ProjName, HorizontalAlignment.Center, VerticalAlignment.Center);
                    for (int i = 5; i < 12; i++)
                    {
                        ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, i, "", HorizontalAlignment.Center,
                            VerticalAlignment.Center);
                    }

                    sheet.AddMergedRegion(new CellRangeAddress(rowIndex, rowIndex, 4, 11));
                    colIndex = 12;
                    ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, colIndex++, iRowSUM.ProjCount,
                        HorizontalAlignment.Center, VerticalAlignment.Center);
                    ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, colIndex++, "",
                        HorizontalAlignment.Center, VerticalAlignment.Center);
                    sheet.AddMergedRegion(new CellRangeAddress(rowIndex, rowIndex, 12, 13));
                    ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, colIndex++, iRowSUM.HoCount,
                        HorizontalAlignment.Center, VerticalAlignment.Center);
                    ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, colIndex++, iRowSUM.HoProjCode,
                        HorizontalAlignment.Center, VerticalAlignment.Center);
                    ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, colIndex++, "",
                        HorizontalAlignment.Center, VerticalAlignment.Center);
                    sheet.AddMergedRegion(new CellRangeAddress(rowIndex, rowIndex, 15, 16));
                    ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, colIndex++, iRowSUM.HoNullCount,
                        HorizontalAlignment.Center, VerticalAlignment.Center);
                    ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, colIndex++, iRowSUM.HoNullProjCode,
                        HorizontalAlignment.Center, VerticalAlignment.Center);
                    ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, colIndex++, "",
                        HorizontalAlignment.Center, VerticalAlignment.Center);
                    sheet.AddMergedRegion(new CellRangeAddress(rowIndex, rowIndex, 18, 19));
                    for (int i = 20; i < 22; i++)
                    {
                        ExcelHelper.CreateDetail(ref workbook, ref sheet, ref row, i, "", true);
                    }

                    rowIndex++;
                    colIndex = 0;

                    #endregion

                    iSum++;
                    if (iRowSUM.ProjName != "")
                    {
                        sProj += Convert.ToInt32(iRowSUM.ProjCount);
                    }

                    if (iRowSUM.HoCount != "")
                    {
                        sHO += Convert.ToInt32(iRowSUM.HoCount);
                    }

                    if (iRowSUM.HoNullCount != "")
                    {
                        sHO_NULL += Convert.ToInt32(iRowSUM.HoNullCount);
                    }

                    #endregion
                }

                #region รวม

                row = sheet.CreateRow(rowIndex);
                for (int i = 0; i < 4; i++)
                {
                    ExcelHelper.CreateDetail(ref workbook, ref sheet, ref row, i, "", true);
                }

                colIndex = 4;
                ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, colIndex++, "รวม",
                    HorizontalAlignment.Center, VerticalAlignment.Center);
                for (int i = 5; i < 12; i++)
                {
                    ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, i, "", HorizontalAlignment.Center,
                        VerticalAlignment.Center);
                }

                sheet.AddMergedRegion(new CellRangeAddress(rowIndex, rowIndex, 4, 11));
                colIndex = 12;
                ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, colIndex++, sProj,
                    HorizontalAlignment.Center, VerticalAlignment.Center);
                ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, colIndex++, "",
                    HorizontalAlignment.Center, VerticalAlignment.Center);
                sheet.AddMergedRegion(new CellRangeAddress(rowIndex, rowIndex, 12, 13));
                ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, colIndex++, sHO,
                    HorizontalAlignment.Center, VerticalAlignment.Center);
                ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, colIndex++, "",
                    HorizontalAlignment.Center, VerticalAlignment.Center);
                ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, colIndex++, "",
                    HorizontalAlignment.Center, VerticalAlignment.Center);
                sheet.AddMergedRegion(new CellRangeAddress(rowIndex, rowIndex, 15, 16));
                ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, colIndex++, sHO_NULL,
                    HorizontalAlignment.Center, VerticalAlignment.Center);
                ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, colIndex++, "",
                    HorizontalAlignment.Center, VerticalAlignment.Center);
                ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, colIndex++, "",
                    HorizontalAlignment.Center, VerticalAlignment.Center);
                sheet.AddMergedRegion(new CellRangeAddress(rowIndex, rowIndex, 18, 19));
                for (int i = 20; i < 22; i++)
                {
                    ExcelHelper.CreateDetail(ref workbook, ref sheet, ref row, i, "", true);
                }

                rowIndex++;
                colIndex = 0;

                #endregion
            }
            else
            {
                row = sheet.CreateRow(rowIndex);
                for (int i = 0; i < 4; i++)
                {
                    ExcelHelper.CreateDetail(ref workbook, ref sheet, ref row, i, "", true);
                }

                colIndex = 4;
                for (int i = 5; i < 19; i++)
                {
                    ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, i, "", HorizontalAlignment.Center,
                        VerticalAlignment.Center);
                }

                sheet.AddMergedRegion(new CellRangeAddress(rowIndex, rowIndex, 4, 11));
                sheet.AddMergedRegion(new CellRangeAddress(rowIndex, rowIndex, 12, 13));
                sheet.AddMergedRegion(new CellRangeAddress(rowIndex, rowIndex, 15, 16));
                sheet.AddMergedRegion(new CellRangeAddress(rowIndex, rowIndex, 18, 19));
                for (int i = 20; i < 22; i++)
                {
                    ExcelHelper.CreateDetail(ref workbook, ref sheet, ref row, i, "", true);
                }

                rowIndex++;
                colIndex = 0;

                #region รวม

                row = sheet.CreateRow(rowIndex);
                for (int i = 0; i < 4; i++)
                {
                    ExcelHelper.CreateDetail(ref workbook, ref sheet, ref row, i, "", true);
                }

                colIndex = 4;
                ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, colIndex++, "รวม",
                    HorizontalAlignment.Center, VerticalAlignment.Center);
                for (int i = 5; i < 12; i++)
                {
                    ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, i, "", HorizontalAlignment.Center,
                        VerticalAlignment.Center);
                }

                sheet.AddMergedRegion(new CellRangeAddress(rowIndex, rowIndex, 4, 11));
                colIndex = 12;
                ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, colIndex++, "",
                    HorizontalAlignment.Center, VerticalAlignment.Center);
                ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, colIndex++, "",
                    HorizontalAlignment.Center, VerticalAlignment.Center);
                sheet.AddMergedRegion(new CellRangeAddress(rowIndex, rowIndex, 12, 13));
                ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, colIndex++, "",
                    HorizontalAlignment.Center, VerticalAlignment.Center);
                ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, colIndex++, "",
                    HorizontalAlignment.Center, VerticalAlignment.Center);
                ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, colIndex++, "",
                    HorizontalAlignment.Center, VerticalAlignment.Center);
                sheet.AddMergedRegion(new CellRangeAddress(rowIndex, rowIndex, 15, 16));
                ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, colIndex++, "",
                    HorizontalAlignment.Center, VerticalAlignment.Center);
                ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, colIndex++, "",
                    HorizontalAlignment.Center, VerticalAlignment.Center);
                ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, colIndex++, "",
                    HorizontalAlignment.Center, VerticalAlignment.Center);
                sheet.AddMergedRegion(new CellRangeAddress(rowIndex, rowIndex, 18, 19));
                for (int i = 20; i < 22; i++)
                {
                    ExcelHelper.CreateDetail(ref workbook, ref sheet, ref row, i, "", true);
                }

                rowIndex++;
                colIndex = 0;

                #endregion
            }

            row = sheet.CreateRow(rowIndex);
            ExcelHelper.CreateDetail(ref workbook, ref sheet, ref row, colIndex++, "", false);
            sheet.AddMergedRegion(new CellRangeAddress(rowIndex, rowIndex, 0, 21));
            rowIndex++;
            colIndex = 0;

            #region กรณีไม่บรรลุวัตถุประสงค์คุณภาพตามที่กำหนด กรุณาวิเคราะห์ปัญหาและกรอกข้อมูลในด้านล่าง

            row = sheet.CreateRow(rowIndex);
            ExcelHelper.CreateDetail(ref workbook, ref sheet, ref row, colIndex++,
                "กรณีไม่บรรลุวัตถุประสงค์คุณภาพตามที่กำหนด กรุณาวิเคราะห์ปัญหาและกรอกข้อมูลในด้านล่าง", true);
            sheet.AddMergedRegion(new CellRangeAddress(rowIndex, rowIndex, 0, 21));
            rowIndex++;
            colIndex = 0;

            #endregion

            #region สาเหตุที่ทำให้ไม่ได้ตามเป้าหมาย :

            row = sheet.CreateRow(rowIndex);
            ExcelHelper.CreateDetail(ref workbook, ref sheet, ref row, colIndex++, "สาเหตุที่ทำให้ไม่ได้ตามเป้าหมาย:",
                true);
            sheet.AddMergedRegion(new CellRangeAddress(rowIndex, rowIndex, 0, 21));
            rowIndex++;
            colIndex = 0;

            #endregion

            row = sheet.CreateRow(rowIndex);
            ExcelHelper.CreateDetail(ref workbook, ref sheet, ref row, colIndex++, "", false);
            sheet.AddMergedRegion(new CellRangeAddress(rowIndex, rowIndex, 0, 21));
            rowIndex++;
            colIndex = 0;
            row = sheet.CreateRow(rowIndex);
            ExcelHelper.CreateDetail(ref workbook, ref sheet, ref row, colIndex++, "", false);
            sheet.AddMergedRegion(new CellRangeAddress(rowIndex, rowIndex, 0, 21));
            rowIndex++;
            colIndex = 0;
            row = sheet.CreateRow(rowIndex);
            ExcelHelper.CreateDetail(ref workbook, ref sheet, ref row, colIndex++, "", false);
            sheet.AddMergedRegion(new CellRangeAddress(rowIndex, rowIndex, 0, 21));
            rowIndex++;
            colIndex = 0;

            #region การแก้ไข/ป้องกันไม่ให้เกิดซ้ำ: (ผู้บังคับบัญชามีหน้าที่ติดตามให้เกิดผลในทางปฏิบัติ)

            row = sheet.CreateRow(rowIndex);
            ExcelHelper.CreateDetail(ref workbook, ref sheet, ref row, colIndex++,
                "การแก้ไข/ป้องกันไม่ให้เกิดซ้ำ: (ผู้บังคับบัญชามีหน้าที่ติดตามให้เกิดผลในทางปฏิบัติ)", false);
            sheet.AddMergedRegion(new CellRangeAddress(rowIndex, rowIndex, 0, 21));
            rowIndex++;
            colIndex = 0;

            #endregion

            #region การแก้ไข

            row = sheet.CreateRow(rowIndex);
            ExcelHelper.CreateDetail(ref workbook, ref sheet, ref row, colIndex++, "", true);
            ExcelHelper.CreateDetail(ref workbook, ref sheet, ref row, colIndex++, "การแก้ไข", false);
            sheet.AddMergedRegion(new CellRangeAddress(rowIndex, rowIndex, 1, 21));
            rowIndex++;
            colIndex = 0;

            #endregion

            row = sheet.CreateRow(rowIndex);
            ExcelHelper.CreateDetail(ref workbook, ref sheet, ref row, colIndex++, "", false);
            sheet.AddMergedRegion(new CellRangeAddress(rowIndex, rowIndex, 0, 21));
            rowIndex++;
            colIndex = 0;
            row = sheet.CreateRow(rowIndex);
            ExcelHelper.CreateDetail(ref workbook, ref sheet, ref row, colIndex++, "", false);
            sheet.AddMergedRegion(new CellRangeAddress(rowIndex, rowIndex, 0, 21));
            rowIndex++;
            colIndex = 0;
            row = sheet.CreateRow(rowIndex);
            ExcelHelper.CreateDetail(ref workbook, ref sheet, ref row, colIndex++, "", false);
            sheet.AddMergedRegion(new CellRangeAddress(rowIndex, rowIndex, 0, 21));
            rowIndex++;
            colIndex = 0;

            #region การป้องกันไม่ให้ปัญหาเกิดซ้ำ

            row = sheet.CreateRow(rowIndex);
            ExcelHelper.CreateDetail(ref workbook, ref sheet, ref row, colIndex++, "", true);
            ExcelHelper.CreateDetail(ref workbook, ref sheet, ref row, colIndex++, "การป้องกันไม่ให้ปัญหาเกิดซ้ำ",
                false);
            sheet.AddMergedRegion(new CellRangeAddress(rowIndex, rowIndex, 1, 21));
            rowIndex++;
            colIndex = 0;

            #endregion

            row = sheet.CreateRow(rowIndex);
            ExcelHelper.CreateDetail(ref workbook, ref sheet, ref row, colIndex++, "", false);
            sheet.AddMergedRegion(new CellRangeAddress(rowIndex, rowIndex, 0, 21));
            rowIndex++;
            colIndex = 0;
            row = sheet.CreateRow(rowIndex);
            ExcelHelper.CreateDetail(ref workbook, ref sheet, ref row, colIndex++, "", false);
            sheet.AddMergedRegion(new CellRangeAddress(rowIndex, rowIndex, 0, 21));
            rowIndex++;
            colIndex = 0;
            row = sheet.CreateRow(rowIndex);
            ExcelHelper.CreateDetail(ref workbook, ref sheet, ref row, colIndex++, "", false);
            sheet.AddMergedRegion(new CellRangeAddress(rowIndex, rowIndex, 0, 21));
            rowIndex++;
            colIndex = 0;

            #region ผู้จัดทำ ผู้บริหารสายงาน

            row = sheet.CreateRow(rowIndex);
            colIndex = 2;
            ExcelHelper.CreateSign(ref workbook, ref sheet, ref row, colIndex++,
                (object) "ผู้จัดทำ ..................................................................................",
                true);
            for (int i = 3; i < 10; i++)
            {
                ExcelHelper.CreateSign(ref workbook, ref sheet, ref row, i, (object) "", false);
            }

            sheet.AddMergedRegion(new CellRangeAddress(rowIndex, rowIndex, 2, 9));
            colIndex = 12;
            ExcelHelper.CreateSign(ref workbook, ref sheet, ref row, colIndex++,
                (object) "ผู้บริหารสายงาน...............................................", true);
            for (int i = 13; i < 20; i++)
            {
                ExcelHelper.CreateSign(ref workbook, ref sheet, ref row, i, (object) "", false);
            }

            sheet.AddMergedRegion(new CellRangeAddress(rowIndex, rowIndex, 12, 19));
            rowIndex++;
            colIndex = 0;

            #endregion

            #region (...)

            row = sheet.CreateRow(rowIndex);
            colIndex = 2;
            ExcelHelper.CreateSign(ref workbook, ref sheet, ref row, colIndex++,
                "(.......................................................................................)", true);
            for (int i = 3; i < 10; i++)
            {
                ExcelHelper.CreateSign(ref workbook, ref sheet, ref row, i, (object) "", false);
            }

            sheet.AddMergedRegion(new CellRangeAddress(rowIndex, rowIndex, 2, 9));
            colIndex = 12;
            ExcelHelper.CreateSign(ref workbook, ref sheet, ref row, colIndex++,
                (object) "(.......................................................................................)",
                true);
            for (int i = 13; i < 20; i++)
            {
                ExcelHelper.CreateSign(ref workbook, ref sheet, ref row, i, (object) "", false);
            }

            sheet.AddMergedRegion(new CellRangeAddress(rowIndex, rowIndex, 12, 19));
            rowIndex++;
            colIndex = 0;

            #endregion

            #region วันที่.....................................

            row = sheet.CreateRow(rowIndex);
            colIndex = 2;
            ExcelHelper.CreateSign(ref workbook, ref sheet, ref row, colIndex++,
                (object) "(วันที่.....................................)", true);
            for (int i = 3; i < 10; i++)
            {
                ExcelHelper.CreateSign(ref workbook, ref sheet, ref row, i, (object) "", false);
            }

            sheet.AddMergedRegion(new CellRangeAddress(rowIndex, rowIndex, 2, 9));
            colIndex = 12;
            ExcelHelper.CreateSign(ref workbook, ref sheet, ref row, colIndex++,
                (object) "(วันที่.....................................)", true);
            for (int i = 13; i < 20; i++)
            {
                ExcelHelper.CreateSign(ref workbook, ref sheet, ref row, i, (object) "", false);
            }

            sheet.AddMergedRegion(new CellRangeAddress(rowIndex, rowIndex, 12, 19));
            rowIndex++;
            colIndex = 0;

            #endregion

            row = sheet.CreateRow(rowIndex);
            ExcelHelper.CreateDetail(ref workbook, ref sheet, ref row, colIndex++, "", false);
            sheet.AddMergedRegion(new CellRangeAddress(rowIndex, rowIndex, 0, 21));
            rowIndex++;
            colIndex = 0;

            #region เรียน รองกรรมการผู้จัดการ และ QMR ความเห็น รองกรรมการผู้จัดการ ความเห็น QMR

            row = sheet.CreateRow(rowIndex);
            ExcelHelper.CreateDetail(ref workbook, ref sheet, ref row, colIndex++, "เรียน รองกรรมการผู้จัดการ และ QMR",
                false);
            for (int i = 1; i < 7; i++)
            {
                ExcelHelper.CreateSign(ref workbook, ref sheet, ref row, i, (object) "", false);
            }

            sheet.AddMergedRegion(new CellRangeAddress(rowIndex, rowIndex, 0, 6));
            colIndex = 8;
            ExcelHelper.CreateComment(ref workbook, ref sheet, ref row, colIndex++, "ความเห็น รองกรรมการผู้จัดการ");
            for (int i = 9; i < 12; i++)
            {
                ExcelHelper.CreateSign(ref workbook, ref sheet, ref row, i, (object) "", false);
            }

            sheet.AddMergedRegion(new CellRangeAddress(rowIndex, rowIndex, 8, 11));
            colIndex = 15;
            ExcelHelper.CreateComment(ref workbook, ref sheet, ref row, colIndex++, "ความเห็น QMR");
            for (int i = 16; i < 22; i++)
            {
                ExcelHelper.CreateSign(ref workbook, ref sheet, ref row, i, (object) "", false);
            }

            sheet.AddMergedRegion(new CellRangeAddress(rowIndex, rowIndex, 15, 21));
            rowIndex++;
            colIndex = 0;

            #endregion

            #region ทราบ อื่น ๆ

            row = sheet.CreateRow(rowIndex);
            colIndex = 8;
            ExcelHelper.CreateComment(ref workbook, ref sheet, ref row, colIndex++, symbol);
            ExcelHelper.CreateDetail(ref workbook, ref sheet, ref row, colIndex++, "ทราบ", false);
            colIndex = 11;
            ExcelHelper.CreateComment(ref workbook, ref sheet, ref row, colIndex++, symbol);
            ExcelHelper.CreateDetail(ref workbook, ref sheet, ref row, colIndex++, "อื่น ๆ", false);
            colIndex = 15;
            ExcelHelper.CreateComment(ref workbook, ref sheet, ref row, colIndex++, "");
            ExcelHelper.CreateDetail(ref workbook, ref sheet, ref row, colIndex++, symbol, false);
            ExcelHelper.CreateDetail(ref workbook, ref sheet, ref row, colIndex++, "ทราบ", false);
            colIndex++;
            ExcelHelper.CreateDetail(ref workbook, ref sheet, ref row, colIndex++, symbol, false);
            ExcelHelper.CreateDetail(ref workbook, ref sheet, ref row, colIndex++, "ทราบ", false);
            rowIndex++;
            colIndex = 0;

            #endregion

            row = sheet.CreateRow(rowIndex);
            colIndex = 8;
            ExcelHelper.CreateComment(ref workbook, ref sheet, ref row, colIndex++, "");
            colIndex = 15;
            ExcelHelper.CreateComment(ref workbook, ref sheet, ref row, colIndex++, "");
            rowIndex++;
            colIndex = 0;
            row = sheet.CreateRow(rowIndex);
            colIndex = 8;
            ExcelHelper.CreateComment(ref workbook, ref sheet, ref row, colIndex++, "");
            colIndex = 15;
            ExcelHelper.CreateComment(ref workbook, ref sheet, ref row, colIndex++, "");
            rowIndex++;
            colIndex = 0;
            row = sheet.CreateRow(rowIndex);
            colIndex = 8;
            ExcelHelper.CreateComment(ref workbook, ref sheet, ref row, colIndex++, "");
            colIndex = 15;
            ExcelHelper.CreateComment(ref workbook, ref sheet, ref row, colIndex++, "");
            rowIndex++;
            colIndex = 0;
            row = sheet.CreateRow(rowIndex);
            colIndex = 8;
            ExcelHelper.CreateComment(ref workbook, ref sheet, ref row, colIndex++, "");
            colIndex = 15;
            ExcelHelper.CreateComment(ref workbook, ref sheet, ref row, colIndex++, "");
            rowIndex++;
            colIndex = 0;
            row = sheet.CreateRow(rowIndex);
            colIndex = 8;
            ExcelHelper.CreateComment(ref workbook, ref sheet, ref row, colIndex++, "");
            colIndex = 15;
            ExcelHelper.CreateComment(ref workbook, ref sheet, ref row, colIndex++, "");
            rowIndex++;
            colIndex = 0;

            #region หน่วยงานระบบคุณภาพ  / วันที่ ลงนาม...

            row = sheet.CreateRow(rowIndex);
            ExcelHelper.CreateDetail(ref workbook, ref sheet, ref row, colIndex++,
                "หน่วยงานระบบคุณภาพ.................................... / วันที่ ..................", false);
            colIndex = 8;
            ExcelHelper.CreateComment(ref workbook, ref sheet, ref row, colIndex++,
                "ลงนาม............................................................................");
            colIndex = 15;
            ExcelHelper.CreateComment(ref workbook, ref sheet, ref row, colIndex++,
                "ลงนาม.......................................................");
            rowIndex++;
            colIndex = 0;

            #endregion

            #region (นางวารุณี ลภิธนานุวัฒน์) (นายกิตติพงษ์ ศิริลักษณ์ตระกูล)

            row = sheet.CreateRow(rowIndex);
            colIndex = 8;
            ExcelHelper.CreateComment(ref workbook, ref sheet, ref row, colIndex++, "(นางวารุณี ลภิธนานุวัฒน์)");
            colIndex = 15;
            ExcelHelper.CreateComment(ref workbook, ref sheet, ref row, colIndex++, "(นายกิตติพงษ์ ศิริลักษณ์ตระกูล)");
            rowIndex++;
            colIndex = 0;

            #endregion

            #region ผู้อำนวยการฝ่ายกำกับดูแลกิจการฯ./ วันที่ .รองกรรมการผู้จัดการ รองกรรมการผู้จัดการ สายงานก่อสร้างอาคารสูง / QMR

            row = sheet.CreateRow(rowIndex);
            ExcelHelper.CreateDetail(ref workbook, ref sheet, ref row, colIndex++,
                "ผู้อำนวยการฝ่ายกำกับดูแลกิจการฯ.................... / วันที่ .................", false);
            colIndex = 8;
            ExcelHelper.CreateComment(ref workbook, ref sheet, ref row, colIndex++, "รองกรรมการผู้จัดการ");
            colIndex = 15;
            ExcelHelper.CreateComment(ref workbook, ref sheet, ref row, colIndex++,
                "รองกรรมการผู้จัดการสายงานก่อสร้างอาคารสูง / QMR");
            rowIndex++;
            colIndex = 0;

            #endregion

            #region วันที่...

            row = sheet.CreateRow(rowIndex);
            colIndex = 8;
            ExcelHelper.CreateComment(ref workbook, ref sheet, ref row, colIndex++,
                "วันที่.......................................................");
            colIndex = 15;
            ExcelHelper.CreateComment(ref workbook, ref sheet, ref row, colIndex++,
                "วันที่.......................................................");
            rowIndex++;
            colIndex = 0;

            #endregion

            row = sheet.CreateRow(rowIndex);
            rowIndex++;
            colIndex = 0;
        }

        public static void CreateKpiHoNullLowRise(ref XSSFWorkbook workbook, ref ISheet sheet, PrmGetRpt prm)
        {
            var valuePeriod = BindingData.setupPeriod(prm);
            var count_ho_null = GetCountKpiHoNullPeriod(prm);
            var count_ho_null_late = GetCountKpiHoNullLateLowRise(prm);

            ExcelHelper.SetupSheet(ref sheet);

            var rowIndex = 0;
            var colIndex = 0;

            IRow row = sheet.CreateRow(rowIndex);
            
            #region add logo
            //String image1 = "assets\\img\\logo-spl.JPG";//Image location address//./images/logo1.png
            string path_img = HttpContext.Current.Server.MapPath("assets\\img\\logo-spl.JPG");//add picture data to this workbook.
            byte[] bytes = File.ReadAllBytes(path_img);
            int pictureIdx = workbook.AddPicture(bytes, PictureType.PNG);

            ICreationHelper helper = workbook.GetCreationHelper();

            // Create the drawing patriarch.  This is the top level container for all shapes.
            IDrawing drawing = sheet.CreateDrawingPatriarch();

            // add a picture shape
            IClientAnchor anchor = helper.CreateClientAnchor();

            //set top-left corner of the picture,
            //subsequent call of Picture#resize() will operate relative to it
            anchor.Col1 = 1;
            anchor.Row1 = 2;

            IPicture pict = drawing.CreatePicture(anchor, pictureIdx);
            pict.Resize(2.5, 2);
            //auto-size picture relative to its top-left corner
            //pict.Resize();
            #endregion

            row = sheet.CreateRow(rowIndex);
            rowIndex++;
            colIndex = 0;

            #region แบบรายงานวัตถุประสงค์คุณภาพ (รูปแบบแนวราบ)

            row = sheet.CreateRow(rowIndex);
            ExcelHelper.createHeader(ref workbook, ref sheet, ref row, colIndex++,
                "แบบรายงานวัตถุประสงค์คุณภาพ (รูปแบบแนวราบ)", HorizontalAlignment.Center, VerticalAlignment.Center);
            sheet.AddMergedRegion(new CellRangeAddress(rowIndex, rowIndex, 0, 21));
            rowIndex++;
            colIndex = 0;

            #endregion

            #region ฝ่ายก่อสร้างแนวราบ

            row = sheet.CreateRow(rowIndex);
            ExcelHelper.createHeader(ref workbook, ref sheet, ref row, colIndex++, "ฝ่ายก่อสร้างแนวราบ",
                HorizontalAlignment.Center, VerticalAlignment.Center);
            sheet.AddMergedRegion(new CellRangeAddress(rowIndex, rowIndex, 0, 21));
            rowIndex++;
            colIndex = 0;

            #endregion

            #region ข้อมูลสิ้นสุด ณ เดือน

            row = sheet.CreateRow(rowIndex);
            ExcelHelper.createHeader(ref workbook, ref sheet, ref row, colIndex++,
                "ข้อมูลสิ้นสุด ณ เดือน " + valuePeriod.monthTH + " " + valuePeriod.yearTH + "",
                HorizontalAlignment.Center, VerticalAlignment.Center);
            sheet.AddMergedRegion(new CellRangeAddress(rowIndex, rowIndex, 0, 21));
            rowIndex++;
            colIndex = 0;

            #endregion

            #region วัตถุประสงค์ด้านคุณภาพที่ : 2   เรื่อง : ต้องส่งมอบบ้านที่คงค้างจากเดือนก่อนให้กับลูกค้าตามแผนงานฝ่ายก่อสร้างแนวราบที่ได้รับอนุมัติจากผู้มีอำนาจ

            row = sheet.CreateRow(rowIndex);
            ExcelHelper.createHeaderDetail(ref workbook, ref sheet, ref row, colIndex++,
                "วัตถุประสงค์ด้านคุณภาพที่ : 2   เรื่อง : ต้องส่งมอบบ้านที่คงค้างจากเดือนก่อนให้กับลูกค้าตามแผนงานฝ่ายก่อสร้างแนวราบที่ได้รับอนุมัติจากผู้มีอำนาจ",
                HorizontalAlignment.Left, VerticalAlignment.Center);
            sheet.AddMergedRegion(new CellRangeAddress(rowIndex, rowIndex, 0, 21));
            rowIndex++;
            colIndex = 0;

            #endregion

            #region เกณฑ์การวัดผล :  100% ในแต่ละเดือน

            row = sheet.CreateRow(rowIndex);
            ExcelHelper.createHeaderDetail(ref workbook, ref sheet, ref row, colIndex++,
                "เกณฑ์การวัดผล :  100% ในแต่ละเดือน", HorizontalAlignment.Left, VerticalAlignment.Center);
            sheet.AddMergedRegion(new CellRangeAddress(rowIndex, rowIndex, 0, 21));
            rowIndex++;
            colIndex = 0;

            #endregion

            #region รายการ

            row = sheet.CreateRow(rowIndex);
            ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, colIndex++, "", HorizontalAlignment.Center,
                VerticalAlignment.Center);
            ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, colIndex++, "รายการ",
                HorizontalAlignment.Center, VerticalAlignment.Center);
            for (int i = 2; i < 9; i++)
            {
                ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, i, "", HorizontalAlignment.Center,
                    VerticalAlignment.Center);
            }

            sheet.AddMergedRegion(new CellRangeAddress(rowIndex, rowIndex, 1, 8));
            colIndex = 9;
            ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, colIndex++, "ม.ค.",
                HorizontalAlignment.Center, VerticalAlignment.Center);
            ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, colIndex++, "ก.พ.",
                HorizontalAlignment.Center, VerticalAlignment.Center);
            ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, colIndex++, "มี.ค.",
                HorizontalAlignment.Center, VerticalAlignment.Center);
            ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, colIndex++, "เม.ย.",
                HorizontalAlignment.Center, VerticalAlignment.Center);
            ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, colIndex++, "พ.ค.",
                HorizontalAlignment.Center, VerticalAlignment.Center);
            ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, colIndex++, "มิ.ย.",
                HorizontalAlignment.Center, VerticalAlignment.Center);
            ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, colIndex++, "ก.ค.",
                HorizontalAlignment.Center, VerticalAlignment.Center);
            ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, colIndex++, "ส.ค.",
                HorizontalAlignment.Center, VerticalAlignment.Center);
            ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, colIndex++, "ก.ย.",
                HorizontalAlignment.Center, VerticalAlignment.Center);
            ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, colIndex++, "ต.ค.",
                HorizontalAlignment.Center, VerticalAlignment.Center);
            ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, colIndex++, "พ.ย.",
                HorizontalAlignment.Center, VerticalAlignment.Center);
            ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, colIndex++, "ธ.ค.",
                HorizontalAlignment.Center, VerticalAlignment.Center);
            ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, colIndex++, "รวม", HorizontalAlignment.Center,
                VerticalAlignment.Center);
            rowIndex++;
            colIndex = 0;

            #endregion

            #region 1.Plan 1. จำนวนแปลงที่คงค้างการส่งมอบจากเดือนก่อนๆ (ยกมาจากข้อ 4 ของ Obj เรื่องที่ 1 ของเดือนก่อน)

            row = sheet.CreateRow(rowIndex);
            ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, colIndex++, "1.Plan",
                HorizontalAlignment.Center, VerticalAlignment.Center);
            ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, colIndex++,
                "1. จำนวนแปลงที่คงค้างการส่งมอบจากเดือนก่อนๆ (ยกมาจากข้อ 4 ของ Obj เรื่องที่ 1 ของเดือนก่อน)",
                HorizontalAlignment.Left, VerticalAlignment.Center);
            for (int i = 2; i < 22; i++)
            {
                ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, i, "", HorizontalAlignment.Center,
                    VerticalAlignment.Center);
            }

            sheet.AddMergedRegion(new CellRangeAddress(rowIndex, rowIndex, 1, 8));

            int sum1 = 0;
            for (int i = 1; i < (valuePeriod.iMonth_current + 1); i++)
            {
                foreach (var item in count_ho_null)
                {
                    if (item.Month == i)
                    {
                        ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, (i + 8), item.PlanHoNullPreMonth,
                            HorizontalAlignment.Center, VerticalAlignment.Center);
                        sum1 += item.PlanHoNullPreMonth;
                    }
                }
            }

            ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, 21, sum1, HorizontalAlignment.Center,
                VerticalAlignment.Center);

            rowIndex++;
            colIndex = 0;

            #endregion

            #region 2.Actual 2. จำนวนแปลงคงค้าง ที่สามารถส่งมอบได้ในเดือน

            row = sheet.CreateRow(rowIndex);
            ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, colIndex++, "2.Actual",
                HorizontalAlignment.Center, VerticalAlignment.Center);
            ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, colIndex++,
                "2. จำนวนแปลงคงค้าง ที่สามารถส่งมอบได้ในเดือน", HorizontalAlignment.Left, VerticalAlignment.Center);
            for (int i = 2; i < 22; i++)
            {
                ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, i, "", HorizontalAlignment.Center,
                    VerticalAlignment.Center);
            }

            sheet.AddMergedRegion(new CellRangeAddress(rowIndex, rowIndex, 1, 8));

            int sum2 = 0;
            for (int i = 1; i < (valuePeriod.iMonth_current + 1); i++)
            {
                foreach (var item in count_ho_null)
                {
                    if (item.Month == i)
                    {
                        ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, (i + 8), item.ActHo,
                            HorizontalAlignment.Center, VerticalAlignment.Center);
                        sum2 += item.ActHo;
                    }
                }
            }

            ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, 21, sum2, HorizontalAlignment.Center,
                VerticalAlignment.Center);

            rowIndex++;
            colIndex = 0;

            #endregion

            #region 3. จำนวนแปลงคงค้างที่ยังไม่สามารถส่งมอบได้ (ข้อ 1 - 2)

            row = sheet.CreateRow(rowIndex);
            ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, colIndex++, "", HorizontalAlignment.Center,
                VerticalAlignment.Center);
            ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, colIndex++,
                "3. จำนวนแปลงคงค้างที่ยังไม่สามารถส่งมอบได้ (ข้อ 1 - 2)", HorizontalAlignment.Left,
                VerticalAlignment.Center);
            for (int i = 2; i < 22; i++)
            {
                ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, i, "", HorizontalAlignment.Center,
                    VerticalAlignment.Center);
            }

            sheet.AddMergedRegion(new CellRangeAddress(rowIndex, rowIndex, 1, 8));

            int sum3 = 0;
            for (int i = 1; i < (valuePeriod.iMonth_current + 1); i++)
            {
                foreach (var item in count_ho_null)
                {
                    if (item.Month == i)
                    {
                        ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, (i + 8), item.ActHoNull,
                            HorizontalAlignment.Center, VerticalAlignment.Center);
                        sum3 += item.ActHoNull;
                    }
                }
            }

            ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, 21, sum3, HorizontalAlignment.Center,
                VerticalAlignment.Center);

            rowIndex++;
            colIndex = 0;

            #endregion

            #region % ที่ส่งมอบบ้านที่คงค้างได้ในเดือน (ข้อ 2 / ข้อ1)

            row = sheet.CreateRow(rowIndex);
            ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, colIndex++, "", HorizontalAlignment.Center,
                VerticalAlignment.Center);
            ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, colIndex++,
                "% ที่ส่งมอบบ้านที่คงค้างได้ในเดือน (ข้อ 2 / ข้อ1)", HorizontalAlignment.Left,
                VerticalAlignment.Center);
            for (int i = 2; i < 22; i++)
            {
                ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, i, "", HorizontalAlignment.Center,
                    VerticalAlignment.Center);
            }

            sheet.AddMergedRegion(new CellRangeAddress(rowIndex, rowIndex, 1, 8));

            double perc = 0;
            int iperc = 0;
            for (int i = 1; i < (valuePeriod.iMonth_current + 1); i++)
            {
                foreach (var item in count_ho_null)
                {
                    if (item.Month == i)
                    {
                        if (item.ActHo == 0)
                        {
                            ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, (i + 8), "-",
                                HorizontalAlignment.Center, VerticalAlignment.Center);
                        }
                        else
                        {
                            ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, (i + 8), item.Perc + "%",
                                HorizontalAlignment.Center, VerticalAlignment.Center);
                            perc += item.Perc;
                            iperc += 1;
                        }
                    }
                }
            }

            perc = (perc / valuePeriod.iMonth_current);
            //sum8 = (float)(Math.Round((double)sum8, 2));
            perc = (float) Math.Round((double) perc);
            if (iperc == 0)
            {
                ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, 21, "-", HorizontalAlignment.Center,
                    VerticalAlignment.Center);
            }
            else
            {
                ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, 21, "-", HorizontalAlignment.Center,
                    VerticalAlignment.Center);
            }

            ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, 21, perc + "%", HorizontalAlignment.Center,
                VerticalAlignment.Center);

            rowIndex++;
            colIndex = 0;

            #endregion

            #region Moniter 4. จำนวนแปลงที่ต้องส่งมอบบ้านคงค้างสะสม ณ เดือนนี้ (ข้อ 3 เดือนปัจจุบัน + ข้อ 6 ของเดือนก่อน)

            row = sheet.CreateRow(rowIndex);
            ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, colIndex++, "Moniter",
                HorizontalAlignment.Center, VerticalAlignment.Center);
            ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, colIndex++,
                "4. จำนวนแปลงที่ต้องส่งมอบบ้านคงค้างสะสม ณ เดือนนี้ (ข้อ 3 เดือนปัจจุบัน + ข้อ 6 ของเดือนก่อน)",
                HorizontalAlignment.Left, VerticalAlignment.Center);
            for (int i = 2; i < 22; i++)
            {
                ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, i, "", HorizontalAlignment.Center,
                    VerticalAlignment.Center);
            }

            sheet.AddMergedRegion(new CellRangeAddress(rowIndex, rowIndex, 1, 8));

            int sum4 = 0;
            for (int i = 1; i < (valuePeriod.iMonth_current + 1); i++)
            {
                foreach (var item in count_ho_null)
                {
                    if (item.Month == i)
                    {
                        ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, (i + 8), item.MoniterHoNull,
                            HorizontalAlignment.Center, VerticalAlignment.Center);
                        sum4 += item.MoniterHoNull;
                    }
                }
            }

            ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, 21, sum4, HorizontalAlignment.Center,
                VerticalAlignment.Center);

            rowIndex++;
            colIndex = 0;

            #endregion

            #region 5. จำนวนแปลงที่คงค้าง จากเดือนก่อนๆ แต่ส่งมอบได้ในเดือนนี้ (คงค้างทุกเดือน รับมอบได้จากข้อ 6 เดือนก่อนหน้า)

            row = sheet.CreateRow(rowIndex);
            ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, colIndex++, "", HorizontalAlignment.Center,
                VerticalAlignment.Center);
            ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, colIndex++,
                "5. จำนวนแปลงที่คงค้าง จากเดือนก่อนๆ แต่ส่งมอบได้ในเดือนนี้ (คงค้างทุกเดือน รับมอบได้จากข้อ 6 เดือนก่อนหน้า)",
                HorizontalAlignment.Left, VerticalAlignment.Center);
            for (int i = 2; i < 22; i++)
            {
                ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, i, "", HorizontalAlignment.Center,
                    VerticalAlignment.Center);
            }

            sheet.AddMergedRegion(new CellRangeAddress(rowIndex, rowIndex, 1, 8));

            int sum5 = 0;
            for (int i = 1; i < (valuePeriod.iMonth_current + 1); i++)
            {
                foreach (var item in count_ho_null)
                {
                    if (item.Month == i)
                    {
                        ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, (i + 8), item.MoniterHo,
                            HorizontalAlignment.Center, VerticalAlignment.Center);
                        sum5 += item.MoniterHo;
                    }
                }
            }

            ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, 21, sum5, HorizontalAlignment.Center,
                VerticalAlignment.Center);

            rowIndex++;
            colIndex = 0;

            #endregion

            #region 6. คงเหลือจำนวนแปลงคงค้างสะสมที่ไม่สามารถส่งมอบบ้านได้ (ข้อ 4 - 5)

            row = sheet.CreateRow(rowIndex);
            ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, colIndex++, "", HorizontalAlignment.Center,
                VerticalAlignment.Center);
            ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, colIndex++,
                "6. คงเหลือจำนวนแปลงคงค้างสะสมที่ไม่สามารถส่งมอบบ้านได้ (ข้อ 4 - 5)", HorizontalAlignment.Left,
                VerticalAlignment.Center);
            for (int i = 2; i < 22; i++)
            {
                ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, i, "", HorizontalAlignment.Center,
                    VerticalAlignment.Center);
            }

            sheet.AddMergedRegion(new CellRangeAddress(rowIndex, rowIndex, 1, 8));

            int sum6 = 0;
            for (int i = 1; i < (valuePeriod.iMonth_current + 1); i++)
            {
                foreach (var item in count_ho_null)
                {
                    if (item.Month == i)
                    {
                        ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, (i + 8), item.HoNull,
                            HorizontalAlignment.Center, VerticalAlignment.Center);
                        sum6 += item.HoNull;
                    }
                }
            }

            ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, 21, sum6, HorizontalAlignment.Center,
                VerticalAlignment.Center);

            rowIndex++;
            colIndex = 0;

            #endregion

            row = sheet.CreateRow(rowIndex);
            rowIndex++;
            colIndex = 0;

            #region หมายเหตุ:

            row = sheet.CreateRow(rowIndex);
            ExcelHelper.CreateDetail(ref workbook, ref sheet, ref row, colIndex++, "หมายเหตุ:", true);
            sheet.AddMergedRegion(new CellRangeAddress(rowIndex, rowIndex, 0, 21));
            rowIndex++;
            colIndex = 0;

            #endregion

            #region แปลงคงค้างจากเดือนก่อน ที่สามารถส่งมอบได้ในเดือนนี้

            row = sheet.CreateRow(rowIndex);
            colIndex = 1;
            ExcelHelper.CreateDetail(ref workbook, ref sheet, ref row, colIndex++,
                "แปลงคงค้างจากเดือนก่อน ที่สามารถส่งมอบได้ในเดือนนี้", false);
            sheet.AddMergedRegion(new CellRangeAddress(rowIndex, rowIndex, 1, 21));
            rowIndex++;
            colIndex = 0;

            #endregion

            row = sheet.CreateRow(rowIndex);
            ExcelHelper.CreateDetail(ref workbook, ref sheet, ref row, colIndex++, "", false);
            sheet.AddMergedRegion(new CellRangeAddress(rowIndex, rowIndex, 0, 21));
            rowIndex++;
            colIndex = 0;
            row = sheet.CreateRow(rowIndex);
            ExcelHelper.CreateDetail(ref workbook, ref sheet, ref row, colIndex++, "", false);
            sheet.AddMergedRegion(new CellRangeAddress(rowIndex, rowIndex, 0, 21));
            rowIndex++;
            colIndex = 0;

            #region แปลงที่คงค้างทั้งหมด(คงค้างส่งมอบสะสม ณ สิ้นที่รายงานผล)

            row = sheet.CreateRow(rowIndex);
            colIndex = 1;
            ExcelHelper.CreateDetail(ref workbook, ref sheet, ref row, colIndex++,
                "แปลงที่คงค้างทั้งหมด(คงค้างส่งมอบสะสม ณ สิ้นที่รายงานผล)", false);
            sheet.AddMergedRegion(new CellRangeAddress(rowIndex, rowIndex, 1, 21));
            rowIndex++;
            colIndex = 0;

            #endregion

            row = sheet.CreateRow(rowIndex);
            rowIndex++;
            colIndex = 0;

            #region แปลงที่คงค้างส่งมอบสะสมทั้งสิ้น ณ สิ้นเดือนที่รายงานผล

            row = sheet.CreateRow(rowIndex);
            colIndex = 4;
            ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, colIndex++,
                "แปลงที่คงค้างส่งมอบสะสมทั้งสิ้น ณ สิ้นเดือนที่รายงานผล", HorizontalAlignment.Center,
                VerticalAlignment.Center);
            for (int i = 5; i < 20; i++)
            {
                ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, i, "", HorizontalAlignment.Center,
                    VerticalAlignment.Center);
            }

            sheet.AddMergedRegion(new CellRangeAddress(rowIndex, rowIndex, 4, 19));
            rowIndex++;
            colIndex = 0;

            #endregion

            #region โครงการ จำนวน ค้างส่งมอบ

            row = sheet.CreateRow(rowIndex);
            colIndex = 4;
            ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, colIndex++, "โครงการ",
                HorizontalAlignment.Center, VerticalAlignment.Center);
            for (int i = 5; i < 12; i++)
            {
                ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, i, "", HorizontalAlignment.Center,
                    VerticalAlignment.Center);
            }

            //โครงการ merged รวม  
            //sheet.AddMergedRegion(new CellRangeAddress(rowIndex, rowIndex, 4, 11));
            colIndex = 12;
            ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, colIndex++, "จำนวน",
                HorizontalAlignment.Center, VerticalAlignment.Center);
            ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, colIndex++, "", HorizontalAlignment.Center,
                VerticalAlignment.Center);
            //จำนวน merged รวม
            //sheet.AddMergedRegion(new CellRangeAddress(rowIndex, rowIndex, 12, 13));
            ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, colIndex++, "ค้างส่งมอบ",
                HorizontalAlignment.Center, VerticalAlignment.Center);
            for (int i = 15; i < 20; i++)
            {
                ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, i, "", HorizontalAlignment.Center,
                    VerticalAlignment.Center);
            }

            sheet.AddMergedRegion(new CellRangeAddress(rowIndex, rowIndex, 14, 19));
            rowIndex++;
            colIndex = 0;

            #endregion

            #region จำนวน เลขที่แปลง

            row = sheet.CreateRow(rowIndex);
            for (int i = 4; i < 12; i++)
            {
                ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, i, "", HorizontalAlignment.Center,
                    VerticalAlignment.Center);
            }

            //โครงการ merged รวม  
            sheet.AddMergedRegion(new CellRangeAddress(rowIndex - 1, rowIndex, 4, 11));
            colIndex = 12;
            ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, colIndex++, "จำนวน",
                HorizontalAlignment.Center, VerticalAlignment.Center);
            ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, colIndex++, "", HorizontalAlignment.Center,
                VerticalAlignment.Center);
            //จำนวน merged รวม
            sheet.AddMergedRegion(new CellRangeAddress(rowIndex - 1, rowIndex, 12, 13));
            colIndex = 14;
            ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, colIndex++, "จำนวน",
                HorizontalAlignment.Center, VerticalAlignment.Center);
            ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, colIndex++, "เลขที่แปลง",
                HorizontalAlignment.Center, VerticalAlignment.Center);
            for (int i = 16; i < 20; i++)
            {
                ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, i, "", HorizontalAlignment.Center,
                    VerticalAlignment.Center);
            }

            sheet.AddMergedRegion(new CellRangeAddress(rowIndex, rowIndex, 15, 19));
            rowIndex++;
            colIndex = 0;

            #endregion

            ///count_ho_null_late
            int iSum = 1;
            int sProj = 0;
            //int sHO = 0;
            int sHO_NULL = 0;

            if (count_ho_null_late.Count > 0)
            {
                foreach (var iRowSUM in count_ho_null_late)
                {
                    #region list detail

                    row = sheet.CreateRow(rowIndex);
                    colIndex = 4;
                    ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, colIndex++,
                        iSum + ". " + iRowSUM.ProjName, HorizontalAlignment.Center, VerticalAlignment.Center);

                    for (int i = 5; i < 12; i++)
                    {
                        ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, i, "", HorizontalAlignment.Center,
                            VerticalAlignment.Center);
                    }

                    ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, 12, iRowSUM.ProjCount,
                        HorizontalAlignment.Center, VerticalAlignment.Center);
                    ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, 13, "", HorizontalAlignment.Center,
                        VerticalAlignment.Center);
                    ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, 14, iRowSUM.HoNullCount,
                        HorizontalAlignment.Center, VerticalAlignment.Center);
                    ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, 15, iRowSUM.HoNullProjCode,
                        HorizontalAlignment.Center, VerticalAlignment.Center);
                    for (int i = 16; i < 20; i++)
                    {
                        ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, i, "", HorizontalAlignment.Center,
                            VerticalAlignment.Center);
                    }

                    sheet.AddMergedRegion(new CellRangeAddress(rowIndex, rowIndex, 4, 11));
                    sheet.AddMergedRegion(new CellRangeAddress(rowIndex, rowIndex, 12, 13));
                    sheet.AddMergedRegion(new CellRangeAddress(rowIndex, rowIndex, 15, 19));

                    rowIndex++;
                    colIndex = 0;

                    iSum++;
                    if (iRowSUM.ProjCount != "")
                    {
                        sProj += Convert.ToInt32(iRowSUM.ProjCount);
                    }

                    if (iRowSUM.HoNullCount != "")
                    {
                        sHO_NULL += Convert.ToInt32(iRowSUM.HoNullCount);
                    }

                    #endregion
                }

                #region รวม

                row = sheet.CreateRow(rowIndex);
                colIndex = 4;
                ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, colIndex++, "รวม",
                    HorizontalAlignment.Center, VerticalAlignment.Center);
                for (int i = 5; i < 12; i++)
                {
                    ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, i, "", HorizontalAlignment.Center,
                        VerticalAlignment.Center);
                }

                ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, 12, sProj, HorizontalAlignment.Center,
                    VerticalAlignment.Center);
                ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, 13, "", HorizontalAlignment.Center,
                    VerticalAlignment.Center);
                ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, 14, sHO_NULL, HorizontalAlignment.Center,
                    VerticalAlignment.Center);
                //ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, 15, "", HorizontalAlignment.Center, VerticalAlignment.Center); 
                for (int i = 15; i < 20; i++)
                {
                    ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, i, "", HorizontalAlignment.Center,
                        VerticalAlignment.Center);
                }

                //โครงการ merged รวม  
                sheet.AddMergedRegion(new CellRangeAddress(rowIndex, rowIndex, 4, 11));
                //จำนวน merged รวม
                sheet.AddMergedRegion(new CellRangeAddress(rowIndex, rowIndex, 12, 13));
                sheet.AddMergedRegion(new CellRangeAddress(rowIndex, rowIndex, 15, 19));
                rowIndex++;
                colIndex = 0;

                #endregion
            }
            else
            {
                row = sheet.CreateRow(rowIndex);
                // colIndex = 4; 
                for (int i = 4; i < 20; i++)
                {
                    ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, i, "", HorizontalAlignment.Center,
                        VerticalAlignment.Center);
                }

                //โครงการ merged รวม  
                sheet.AddMergedRegion(new CellRangeAddress(rowIndex, rowIndex, 4, 11));
                //จำนวน merged รวม
                sheet.AddMergedRegion(new CellRangeAddress(rowIndex, rowIndex, 12, 13));
                sheet.AddMergedRegion(new CellRangeAddress(rowIndex, rowIndex, 15, 19));
                rowIndex++;
                colIndex = 0;

                #region รวม

                row = sheet.CreateRow(rowIndex);
                colIndex = 4;
                ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, colIndex++, "รวม",
                    HorizontalAlignment.Center, VerticalAlignment.Center);
                for (int i = 5; i < 20; i++)
                {
                    ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, i, "", HorizontalAlignment.Center,
                        VerticalAlignment.Center);
                }

                //โครงการ merged รวม  
                sheet.AddMergedRegion(new CellRangeAddress(rowIndex, rowIndex, 4, 11));
                //จำนวน merged รวม
                sheet.AddMergedRegion(new CellRangeAddress(rowIndex, rowIndex, 12, 13));
                sheet.AddMergedRegion(new CellRangeAddress(rowIndex, rowIndex, 15, 19));
                rowIndex++;
                colIndex = 0;

                #endregion
            }

            //#region รวม
            //row = sheet.CreateRow(rowIndex);
            //colIndex = 4;
            //ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, colIndex++, "รวม", HorizontalAlignment.Center, VerticalAlignment.Center);
            //for (int i = 5; i < 20; i++)
            //{
            //    ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, i, "", HorizontalAlignment.Center, VerticalAlignment.Center);
            //}
            ////โครงการ merged รวม  
            //sheet.AddMergedRegion(new CellRangeAddress(rowIndex, rowIndex, 4, 11));
            ////จำนวน merged รวม
            //sheet.AddMergedRegion(new CellRangeAddress(rowIndex, rowIndex, 12, 13));
            //sheet.AddMergedRegion(new CellRangeAddress(rowIndex, rowIndex, 15, 19));
            //rowIndex++; colIndex = 0;
            //#endregion

            #region สาเหตุที่ทำให้ไม่ได้ตามเป้าหมาย:

            row = sheet.CreateRow(rowIndex);
            ExcelHelper.CreateDetail(ref workbook, ref sheet, ref row, colIndex++, "สาเหตุที่ทำให้ไม่ได้ตามเป้าหมาย:",
                true);
            sheet.AddMergedRegion(new CellRangeAddress(rowIndex, rowIndex, 0, 21));
            rowIndex++;
            colIndex = 0;

            #endregion

            row = sheet.CreateRow(rowIndex);
            ExcelHelper.CreateDetail(ref workbook, ref sheet, ref row, colIndex++, "", false);
            sheet.AddMergedRegion(new CellRangeAddress(rowIndex, rowIndex, 0, 21));
            rowIndex++;
            colIndex = 0;
            row = sheet.CreateRow(rowIndex);
            ExcelHelper.CreateDetail(ref workbook, ref sheet, ref row, colIndex++, "", false);
            sheet.AddMergedRegion(new CellRangeAddress(rowIndex, rowIndex, 0, 21));
            rowIndex++;
            colIndex = 0;
            row = sheet.CreateRow(rowIndex);
            ExcelHelper.CreateDetail(ref workbook, ref sheet, ref row, colIndex++, "", false);
            sheet.AddMergedRegion(new CellRangeAddress(rowIndex, rowIndex, 0, 21));
            rowIndex++;
            colIndex = 0;

            #region การแก้ไขปัญหา/ป้องกันไม่ให้เกิดซ้ำ: (ผู้บังคับบัญชามีหน้าที่ติดตามให้เกิดผลในทางปฏิบัติ)

            row = sheet.CreateRow(rowIndex);
            ExcelHelper.CreateDetail(ref workbook, ref sheet, ref row, colIndex++,
                "การแก้ไขปัญหา/ป้องกันไม่ให้เกิดซ้ำ: (ผู้บังคับบัญชามีหน้าที่ติดตามให้เกิดผลในทางปฏิบัติ)", true);
            sheet.AddMergedRegion(new CellRangeAddress(rowIndex, rowIndex, 0, 21));
            rowIndex++;
            colIndex = 0;

            #endregion

            row = sheet.CreateRow(rowIndex);
            ExcelHelper.CreateDetail(ref workbook, ref sheet, ref row, colIndex++, "", false);
            sheet.AddMergedRegion(new CellRangeAddress(rowIndex, rowIndex, 0, 21));
            rowIndex++;
            colIndex = 0;
            row = sheet.CreateRow(rowIndex);
            ExcelHelper.CreateDetail(ref workbook, ref sheet, ref row, colIndex++, "", false);
            sheet.AddMergedRegion(new CellRangeAddress(rowIndex, rowIndex, 0, 21));
            rowIndex++;
            colIndex = 0;
            row = sheet.CreateRow(rowIndex);
            ExcelHelper.CreateDetail(ref workbook, ref sheet, ref row, colIndex++, "", false);
            sheet.AddMergedRegion(new CellRangeAddress(rowIndex, rowIndex, 0, 21));
            rowIndex++;
            colIndex = 0;

            #region ผู้จัดทำ ผู้บริหารสายงาน

            row = sheet.CreateRow(rowIndex);
            colIndex = 2;
            ExcelHelper.CreateSign(ref workbook, ref sheet, ref row, colIndex++,
                (object) "ผู้จัดทำ ..................................................................................",
                true);
            for (int i = 3; i < 10; i++)
            {
                ExcelHelper.CreateSign(ref workbook, ref sheet, ref row, i, (object) "", false);
            }

            sheet.AddMergedRegion(new CellRangeAddress(rowIndex, rowIndex, 2, 9));
            colIndex = 12;
            ExcelHelper.CreateSign(ref workbook, ref sheet, ref row, colIndex++,
                (object) "ผู้บริหารสายงาน...............................................", true);
            for (int i = 13; i < 20; i++)
            {
                ExcelHelper.CreateSign(ref workbook, ref sheet, ref row, i, (object) "", false);
            }

            sheet.AddMergedRegion(new CellRangeAddress(rowIndex, rowIndex, 12, 19));
            rowIndex++;
            colIndex = 0;

            #endregion

            #region (...)

            row = sheet.CreateRow(rowIndex);
            colIndex = 2;
            ExcelHelper.CreateSign(ref workbook, ref sheet, ref row, colIndex++,
                (object) "(.......................................................................................)",
                true);
            for (int i = 3; i < 10; i++)
            {
                ExcelHelper.CreateSign(ref workbook, ref sheet, ref row, i, (object) "", false);
            }

            sheet.AddMergedRegion(new CellRangeAddress(rowIndex, rowIndex, 2, 9));
            colIndex = 12;
            ExcelHelper.CreateSign(ref workbook, ref sheet, ref row, colIndex++,
                (object) "(.......................................................................................)",
                true);
            for (int i = 13; i < 20; i++)
            {
                ExcelHelper.CreateSign(ref workbook, ref sheet, ref row, i, (object) "", false);
            }

            sheet.AddMergedRegion(new CellRangeAddress(rowIndex, rowIndex, 12, 19));
            rowIndex++;
            colIndex = 0;

            #endregion

            #region วันที่...

            row = sheet.CreateRow(rowIndex);
            colIndex = 2;
            ExcelHelper.CreateSign(ref workbook, ref sheet, ref row, colIndex++,
                (object) "(วันที่.....................................)", true);
            for (int i = 3; i < 10; i++)
            {
                ExcelHelper.CreateSign(ref workbook, ref sheet, ref row, i, (object) "", false);
            }

            sheet.AddMergedRegion(new CellRangeAddress(rowIndex, rowIndex, 2, 9));
            colIndex = 12;
            ExcelHelper.CreateSign(ref workbook, ref sheet, ref row, colIndex++,
                (object) "(วันที่.....................................)", true);
            for (int i = 13; i < 20; i++)
            {
                ExcelHelper.CreateSign(ref workbook, ref sheet, ref row, i, (object) "", false);
            }

            sheet.AddMergedRegion(new CellRangeAddress(rowIndex, rowIndex, 12, 19));
            rowIndex++;
            colIndex = 0;

            #endregion

            row = sheet.CreateRow(rowIndex);
            ExcelHelper.CreateDetail(ref workbook, ref sheet, ref row, colIndex++, "", false);
            sheet.AddMergedRegion(new CellRangeAddress(rowIndex, rowIndex, 0, 21));
            rowIndex++;
            colIndex = 0;

            #region เรียน รองกรรมการผู้จัดการ และ QMR ความเห็น รองกรรมการผู้จัดการ ความเห็น QMR

            row = sheet.CreateRow(rowIndex);
            ExcelHelper.CreateDetail(ref workbook, ref sheet, ref row, colIndex++, "เรียน รองกรรมการผู้จัดการ และ QMR",
                false);
            for (int i = 1; i < 7; i++)
            {
                ExcelHelper.CreateSign(ref workbook, ref sheet, ref row, i, (object) "", false);
            }

            sheet.AddMergedRegion(new CellRangeAddress(rowIndex, rowIndex, 0, 6));
            colIndex = 8;
            ExcelHelper.CreateComment(ref workbook, ref sheet, ref row, colIndex++, "ความเห็น รองกรรมการผู้จัดการ");
            for (int i = 9; i < 12; i++)
            {
                ExcelHelper.CreateSign(ref workbook, ref sheet, ref row, i, (object) "", false);
            }

            sheet.AddMergedRegion(new CellRangeAddress(rowIndex, rowIndex, 8, 11));
            colIndex = 15;
            ExcelHelper.CreateComment(ref workbook, ref sheet, ref row, colIndex++, "ความเห็น QMR");
            for (int i = 16; i < 22; i++)
            {
                ExcelHelper.CreateSign(ref workbook, ref sheet, ref row, i, (object) "", false);
            }

            sheet.AddMergedRegion(new CellRangeAddress(rowIndex, rowIndex, 15, 21));
            rowIndex++;
            colIndex = 0;

            #endregion

            #region ทราบ อื่น ๆ

            row = sheet.CreateRow(rowIndex);
            colIndex = 8;
            ExcelHelper.CreateComment(ref workbook, ref sheet, ref row, colIndex++, symbol);
            ExcelHelper.CreateDetail(ref workbook, ref sheet, ref row, colIndex++, "ทราบ", false);
            colIndex = 11;
            ExcelHelper.CreateComment(ref workbook, ref sheet, ref row, colIndex++, symbol);
            ExcelHelper.CreateDetail(ref workbook, ref sheet, ref row, colIndex++, "อื่น ๆ", false);
            colIndex = 15;
            ExcelHelper.CreateComment(ref workbook, ref sheet, ref row, colIndex++, "");
            ExcelHelper.CreateDetail(ref workbook, ref sheet, ref row, colIndex++, symbol, false);
            ExcelHelper.CreateDetail(ref workbook, ref sheet, ref row, colIndex++, "ทราบ", false);
            colIndex++;
            ExcelHelper.CreateDetail(ref workbook, ref sheet, ref row, colIndex++, symbol, false);
            ExcelHelper.CreateDetail(ref workbook, ref sheet, ref row, colIndex++, "ทราบ", false);
            rowIndex++;
            colIndex = 0;

            #endregion

            row = sheet.CreateRow(rowIndex);
            colIndex = 8;
            ExcelHelper.CreateComment(ref workbook, ref sheet, ref row, colIndex++, "");
            colIndex = 15;
            ExcelHelper.CreateComment(ref workbook, ref sheet, ref row, colIndex++, "");
            rowIndex++;
            colIndex = 0;
            row = sheet.CreateRow(rowIndex);
            colIndex = 8;
            ExcelHelper.CreateComment(ref workbook, ref sheet, ref row, colIndex++, "");
            colIndex = 15;
            ExcelHelper.CreateComment(ref workbook, ref sheet, ref row, colIndex++, "");
            rowIndex++;
            colIndex = 0;
            row = sheet.CreateRow(rowIndex);
            colIndex = 8;
            ExcelHelper.CreateComment(ref workbook, ref sheet, ref row, colIndex++, "");
            colIndex = 15;
            ExcelHelper.CreateComment(ref workbook, ref sheet, ref row, colIndex++, "");
            rowIndex++;
            colIndex = 0;
            row = sheet.CreateRow(rowIndex);
            colIndex = 8;
            ExcelHelper.CreateComment(ref workbook, ref sheet, ref row, colIndex++, "");
            colIndex = 15;
            ExcelHelper.CreateComment(ref workbook, ref sheet, ref row, colIndex++, "");
            rowIndex++;
            colIndex = 0;
            row = sheet.CreateRow(rowIndex);
            colIndex = 8;
            ExcelHelper.CreateComment(ref workbook, ref sheet, ref row, colIndex++, "");
            colIndex = 15;
            ExcelHelper.CreateComment(ref workbook, ref sheet, ref row, colIndex++, "");
            rowIndex++;
            colIndex = 0;

            #region หน่วยงานระบบคุณภาพ  / วันที่ ลงนาม...

            row = sheet.CreateRow(rowIndex);
            ExcelHelper.CreateDetail(ref workbook, ref sheet, ref row, colIndex++,
                "หน่วยงานระบบคุณภาพ.................................... / วันที่ ..................", false);
            colIndex = 8;
            ExcelHelper.CreateComment(ref workbook, ref sheet, ref row, colIndex++,
                "ลงนาม............................................................................");
            colIndex = 15;
            ExcelHelper.CreateComment(ref workbook, ref sheet, ref row, colIndex++,
                "ลงนาม.......................................................");
            rowIndex++;
            colIndex = 0;

            #endregion

            #region (นางวารุณี ลภิธนานุวัฒน์) (นายกิตติพงษ์ ศิริลักษณ์ตระกูล)

            row = sheet.CreateRow(rowIndex);
            colIndex = 8;
            ExcelHelper.CreateComment(ref workbook, ref sheet, ref row, colIndex++, "(นางวารุณี ลภิธนานุวัฒน์)");
            colIndex = 15;
            ExcelHelper.CreateComment(ref workbook, ref sheet, ref row, colIndex++, "(นายกิตติพงษ์ ศิริลักษณ์ตระกูล)");
            rowIndex++;
            colIndex = 0;

            #endregion

            #region ผู้อำนวยการฝ่ายกำกับดูแลกิจการฯ./ วันที่ .รองกรรมการผู้จัดการ รองกรรมการผู้จัดการ สายงานก่อสร้างอาคารสูง / QMR

            row = sheet.CreateRow(rowIndex);
            ExcelHelper.CreateDetail(ref workbook, ref sheet, ref row, colIndex++,
                "ผู้อำนวยการฝ่ายกำกับดูแลกิจการฯ.................... / วันที่ .................", false);
            colIndex = 8;
            ExcelHelper.CreateComment(ref workbook, ref sheet, ref row, colIndex++, "รองกรรมการผู้จัดการ");
            colIndex = 15;
            ExcelHelper.CreateComment(ref workbook, ref sheet, ref row, colIndex++,
                "รองกรรมการผู้จัดการสายงานก่อสร้างอาคารสูง / QMR");
            rowIndex++;
            colIndex = 0;

            #endregion

            #region วันที่...

            row = sheet.CreateRow(rowIndex);
            colIndex = 8;
            ExcelHelper.CreateComment(ref workbook, ref sheet, ref row, colIndex++,
                "วันที่.......................................................");
            colIndex = 15;
            ExcelHelper.CreateComment(ref workbook, ref sheet, ref row, colIndex++,
                "วันที่.......................................................");
            rowIndex++;
            colIndex = 0;

            #endregion

            row = sheet.CreateRow(rowIndex);
            rowIndex++;
            colIndex = 0;
        }

        public static void CreateKpiFixLowRise(ref XSSFWorkbook workbook, ref ISheet sheet, PrmGetRpt prm)
        {
            var valuePeriod = BindingData.setupPeriod(prm);
            var count_fix = GetCountKpiFixLowRise(prm);

            ExcelHelper.SetupSheet(ref sheet);

            var rowIndex = 0;
            var colIndex = 0;

            IRow row = sheet.CreateRow(rowIndex);
            
            #region add logo
            //String image1 = "assets\\img\\logo-spl.JPG";//Image location address//./images/logo1.png
            string path_img = HttpContext.Current.Server.MapPath("assets\\img\\logo-spl.JPG");//add picture data to this workbook.
            byte[] bytes = File.ReadAllBytes(path_img);
            int pictureIdx = workbook.AddPicture(bytes, PictureType.PNG);

            ICreationHelper helper = workbook.GetCreationHelper();

            // Create the drawing patriarch.  This is the top level container for all shapes.
            IDrawing drawing = sheet.CreateDrawingPatriarch();

            // add a picture shape
            IClientAnchor anchor = helper.CreateClientAnchor();

            //set top-left corner of the picture,
            //subsequent call of Picture#resize() will operate relative to it
            anchor.Col1 = 1;
            anchor.Row1 = 2;

            IPicture pict = drawing.CreatePicture(anchor, pictureIdx);
            pict.Resize(2.5, 2);
            //auto-size picture relative to its top-left corner
            //pict.Resize();
            #endregion

            row = sheet.CreateRow(rowIndex);
            rowIndex++;
            colIndex = 0;

            #region แบบรายงานวัตถุประสงค์คุณภาพ (รูปแบบแนวราบ)

            row = sheet.CreateRow(rowIndex);
            ExcelHelper.createHeader(ref workbook, ref sheet, ref row, colIndex++,
                "แบบรายงานวัตถุประสงค์คุณภาพ (รูปแบบแนวราบ)", HorizontalAlignment.Center, VerticalAlignment.Center);
            sheet.AddMergedRegion(new CellRangeAddress(rowIndex, rowIndex, 0, 21));
            rowIndex++;
            colIndex = 0;

            #endregion

            #region ฝ่ายก่อสร้างแนวราบ

            row = sheet.CreateRow(rowIndex);
            ExcelHelper.createHeader(ref workbook, ref sheet, ref row, colIndex++, "ฝ่ายก่อสร้างแนวราบ",
                HorizontalAlignment.Center, VerticalAlignment.Center);
            sheet.AddMergedRegion(new CellRangeAddress(rowIndex, rowIndex, 0, 21));
            rowIndex++;
            colIndex = 0;

            #endregion

            #region ข้อมูลสิ้นสุด ณ เดือน

            row = sheet.CreateRow(rowIndex);
            ExcelHelper.createHeader(ref workbook, ref sheet, ref row, colIndex++,
                "ข้อมูลสิ้นสุด ณ เดือน " + valuePeriod.monthTH + " " + valuePeriod.yearTH + "",
                HorizontalAlignment.Center, VerticalAlignment.Center);
            sheet.AddMergedRegion(new CellRangeAddress(rowIndex, rowIndex, 0, 21));
            rowIndex++;
            colIndex = 0;

            #endregion

            #region วัตถุประสงค์ด้านคุณภาพที่ : 3   เรื่อง : ซ่อมงานให้ลูกค้าในประกันให้แล้วเสร็จ ภายในกำหนดเวลาที่ตกลงกับลูกค้า

            row = sheet.CreateRow(rowIndex);
            ExcelHelper.createHeaderDetail(ref workbook, ref sheet, ref row, colIndex++,
                "วัตถุประสงค์ด้านคุณภาพที่ : 3   เรื่อง : ซ่อมงานให้ลูกค้าในประกันให้แล้วเสร็จ ภายในกำหนดเวลาที่ตกลงกับลูกค้า",
                HorizontalAlignment.Left, VerticalAlignment.Center);
            sheet.AddMergedRegion(new CellRangeAddress(rowIndex, rowIndex, 0, 21));
            rowIndex++;
            colIndex = 0;

            #endregion

            #region เกณฑ์การวัดผล :  ต้องได้  ≥ 99% ในแต่ละเดือน

            row = sheet.CreateRow(rowIndex);
            ExcelHelper.createHeaderDetail(ref workbook, ref sheet, ref row, colIndex++,
                "เกณฑ์การวัดผล :  ต้องได้  ≥ 99% ในแต่ละเดือน", HorizontalAlignment.Left, VerticalAlignment.Center);
            sheet.AddMergedRegion(new CellRangeAddress(rowIndex, rowIndex, 0, 21));
            rowIndex++;
            colIndex = 0;

            #endregion

            #region รายการ

            row = sheet.CreateRow(rowIndex);
            ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, colIndex++, "", HorizontalAlignment.Center,
                VerticalAlignment.Center);
            ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, colIndex++, "รายการ",
                HorizontalAlignment.Center, VerticalAlignment.Center);
            for (int i = 2; i < 9; i++)
            {
                ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, i, "", HorizontalAlignment.Center,
                    VerticalAlignment.Center);
            }

            sheet.AddMergedRegion(new CellRangeAddress(rowIndex, rowIndex, 1, 8));
            colIndex = 9;
            ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, colIndex++, "ม.ค.",
                HorizontalAlignment.Center, VerticalAlignment.Center);
            ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, colIndex++, "ก.พ.",
                HorizontalAlignment.Center, VerticalAlignment.Center);
            ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, colIndex++, "มี.ค.",
                HorizontalAlignment.Center, VerticalAlignment.Center);
            ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, colIndex++, "เม.ย.",
                HorizontalAlignment.Center, VerticalAlignment.Center);
            ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, colIndex++, "พ.ค.",
                HorizontalAlignment.Center, VerticalAlignment.Center);
            ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, colIndex++, "มิ.ย.",
                HorizontalAlignment.Center, VerticalAlignment.Center);
            ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, colIndex++, "ก.ค.",
                HorizontalAlignment.Center, VerticalAlignment.Center);
            ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, colIndex++, "ส.ค.",
                HorizontalAlignment.Center, VerticalAlignment.Center);
            ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, colIndex++, "ก.ย.",
                HorizontalAlignment.Center, VerticalAlignment.Center);
            ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, colIndex++, "ต.ค.",
                HorizontalAlignment.Center, VerticalAlignment.Center);
            ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, colIndex++, "พ.ย.",
                HorizontalAlignment.Center, VerticalAlignment.Center);
            ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, colIndex++, "ธ.ค.",
                HorizontalAlignment.Center, VerticalAlignment.Center);
            ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, colIndex++, "รวม", HorizontalAlignment.Center,
                VerticalAlignment.Center);
            rowIndex++;
            colIndex = 0;

            #endregion

            #region 1. Plan 1.1 จำนวนใบแจ้งซ่อมที่กำหนดเสร็จในเดือนที่รายงาน (ไม่นับรวมใบแจ้งซ่อมที่ปิดงานและรายงานผลแล้วเมื่อเดือนก่อนๆ)

            row = sheet.CreateRow(rowIndex);
            ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, colIndex++, "1.Plan",
                HorizontalAlignment.Center, VerticalAlignment.Center);
            ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, colIndex++,
                "1.1 จำนวนใบแจ้งซ่อมที่กำหนดเสร็จในเดือนที่รายงาน (ไม่นับรวมใบแจ้งซ่อมที่ปิดงานและรายงานผลแล้วเมื่อเดือนก่อนๆ)",
                HorizontalAlignment.Left, VerticalAlignment.Center);
            for (int i = 2; i < 22; i++)
            {
                ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, i, "", HorizontalAlignment.Center,
                    VerticalAlignment.Center);
            }

            sheet.AddMergedRegion(new CellRangeAddress(rowIndex, rowIndex, 1, 8));

            int sum1 = 0;
            for (int i = 1; i < (valuePeriod.iMonth_current + 1); i++)
            {
                foreach (var item in count_fix)
                {
                    if (item.Month == i)
                    {
                        ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, (i + 8),
                            item.PlanFixDuedateInMounth, HorizontalAlignment.Center,
                            VerticalAlignment.Center);
                        sum1 += item.PlanFixDuedateInMounth;
                    }
                }
            }

            ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, 21, sum1, HorizontalAlignment.Center,
                VerticalAlignment.Center);

            rowIndex++;
            colIndex = 0;

            #endregion

            #region 1.2 จำนวนใบแจ้งซ่อมที่มีกำหนดแล้วเสร็จเดือนต่อๆ ไป แต่สามารถปิดงานได้ก่อนในเดือนที่รายงาน

            row = sheet.CreateRow(rowIndex);
            ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, colIndex++, "", HorizontalAlignment.Center,
                VerticalAlignment.Center);
            ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, colIndex++,
                "1.2 จำนวนใบแจ้งซ่อมที่มีกำหนดแล้วเสร็จเดือนต่อๆ ไป แต่สามารถปิดงานได้ก่อนในเดือนที่รายงาน",
                HorizontalAlignment.Left, VerticalAlignment.Center);
            for (int i = 2; i < 22; i++)
            {
                ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, i, "", HorizontalAlignment.Center,
                    VerticalAlignment.Center);
            }

            sheet.AddMergedRegion(new CellRangeAddress(rowIndex, rowIndex, 1, 8));

            int sum2 = 0;
            for (int i = 1; i < (valuePeriod.iMonth_current + 1); i++)
            {
                foreach (var item in count_fix)
                {
                    if (item.Month == i)
                    {
                        ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, (i + 8),
                            item.PlanFixCloseInMonth, HorizontalAlignment.Center,
                            VerticalAlignment.Center);
                        sum2 += item.PlanFixCloseInMonth;
                    }
                }
            }

            ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, 21, sum2, HorizontalAlignment.Center,
                VerticalAlignment.Center);

            rowIndex++;
            colIndex = 0;

            #endregion

            #region 1.3 จำนวนใบแจ้งซ่อมที่คงค้างการแก้ไขจากเดือนก่อน (ยกมาจาก \"ข้อ 4\" ของเดือนก่อนๆ)

            row = sheet.CreateRow(rowIndex);
            ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, colIndex++, "", HorizontalAlignment.Center,
                VerticalAlignment.Center);
            ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, colIndex++,
                "1.3 จำนวนใบแจ้งซ่อมที่คงค้างการแก้ไขจากเดือนก่อน (ยกมาจาก \"ข้อ 4\" ของเดือนก่อนๆ)",
                HorizontalAlignment.Left, VerticalAlignment.Center);
            for (int i = 2; i < 22; i++)
            {
                ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, i, "", HorizontalAlignment.Center,
                    VerticalAlignment.Center);
            }

            sheet.AddMergedRegion(new CellRangeAddress(rowIndex, rowIndex, 1, 8));

            int sum3 = 0;
            for (int i = 1; i < (valuePeriod.iMonth_current + 1); i++)
            {
                foreach (var item in count_fix)
                {
                    if (item.Month == i)
                    {
                        ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, (i + 8),
                            item.PlanFixHo, HorizontalAlignment.Center,
                            VerticalAlignment.Center);
                        sum3 += item.PlanFixHo;
                    }
                }
            }

            ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, 21, sum3, HorizontalAlignment.Center,
                VerticalAlignment.Center);

            rowIndex++;
            colIndex = 0;

            #endregion

            #region จำนวนใบแจ้งซ่อมที่ต้องแก้ไข้แล้วเสร็จทั้งสิ้นภายในเดือนที่วัดผล (ข้อ 1.1 + ข้อ 1.2 + ข้อ 1.3)

            row = sheet.CreateRow(rowIndex);
            ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, colIndex++, "", HorizontalAlignment.Center,
                VerticalAlignment.Center);
            ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, colIndex++,
                "จำนวนใบแจ้งซ่อมที่ต้องแก้ไข้แล้วเสร็จทั้งสิ้นภายในเดือนที่วัดผล (ข้อ 1.1 + ข้อ 1.2 + ข้อ 1.3)",
                HorizontalAlignment.Left, VerticalAlignment.Center);
            for (int i = 2; i < 22; i++)
            {
                ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, i, "", HorizontalAlignment.Center,
                    VerticalAlignment.Center);
            }

            sheet.AddMergedRegion(new CellRangeAddress(rowIndex, rowIndex, 1, 8));

            int sum4 = 0;
            for (int i = 1; i < (valuePeriod.iMonth_current + 1); i++)
            {
                foreach (var item in count_fix)
                {
                    if (item.Month == i)
                    {
                        ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, (i + 8),
                            item.SumPlan, HorizontalAlignment.Center,
                            VerticalAlignment.Center);
                        sum4 += item.SumPlan;
                    }
                }
            }

            ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, 21, sum4, HorizontalAlignment.Center,
                VerticalAlignment.Center);

            rowIndex++;
            colIndex = 0;

            #endregion

            #region 2. Actual 2.1 ใบแจ้งซ่อมที่มีกำหนดแล้วเสร็จในเดือนที่รายงานที่สามารถปิดได้

            row = sheet.CreateRow(rowIndex);
            ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, colIndex++, "2. Actual",
                HorizontalAlignment.Center, VerticalAlignment.Center);
            ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, colIndex++,
                "2.1 ใบแจ้งซ่อมที่มีกำหนดแล้วเสร็จในเดือนที่รายงานที่สามารถปิดได้", HorizontalAlignment.Left,
                VerticalAlignment.Center);
            for (int i = 2; i < 22; i++)
            {
                ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, i, "", HorizontalAlignment.Center,
                    VerticalAlignment.Center);
            }

            sheet.AddMergedRegion(new CellRangeAddress(rowIndex, rowIndex, 1, 8));

            int sum5 = 0;
            for (int i = 1; i < (valuePeriod.iMonth_current + 1); i++)
            {
                foreach (var item in count_fix)
                {
                    if (item.Month == i)
                    {
                        ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, (i + 8),
                            item.ActualDuedateInMounth, HorizontalAlignment.Center,
                            VerticalAlignment.Center);
                        sum5 += item.ActualDuedateInMounth;
                    }
                }
            }

            ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, 21, sum5, HorizontalAlignment.Center,
                VerticalAlignment.Center);

            rowIndex++;
            colIndex = 0;

            #endregion

            #region 2.2 ใบแจ้งซ่อมที่มีกำหนดเสร็จในเดือนต่อๆ ไป แต่สามารถปิดได้ในเดือนที่รายงานผล

            row = sheet.CreateRow(rowIndex);
            ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, colIndex++, "", HorizontalAlignment.Center,
                VerticalAlignment.Center);
            ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, colIndex++,
                "2.2 ใบแจ้งซ่อมที่มีกำหนดเสร็จในเดือนต่อๆ ไป แต่สามารถปิดได้ในเดือนที่รายงานผล",
                HorizontalAlignment.Left, VerticalAlignment.Center);
            for (int i = 2; i < 22; i++)
            {
                ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, i, "", HorizontalAlignment.Center,
                    VerticalAlignment.Center);
            }

            sheet.AddMergedRegion(new CellRangeAddress(rowIndex, rowIndex, 1, 8));

            int sum6 = 0;
            for (int i = 1; i < (valuePeriod.iMonth_current + 1); i++)
            {
                foreach (var item in count_fix)
                {
                    if (item.Month == i)
                    {
                        ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, (i + 8),
                            item.ActualCloseInMonth, HorizontalAlignment.Center,
                            VerticalAlignment.Center);
                        sum6 += item.ActualCloseInMonth;
                    }
                }
            }

            ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, 21, sum6, HorizontalAlignment.Center,
                VerticalAlignment.Center);

            rowIndex++;
            colIndex = 0;

            #endregion

            #region 2.3 จำนวนใบแจ้งซ่อมที่คงค้างการแก้ไขจากเดือนก่อนๆ แต่ปิดได้ในเดือนนี้

            row = sheet.CreateRow(rowIndex);
            ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, colIndex++, "", HorizontalAlignment.Center,
                VerticalAlignment.Center);
            ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, colIndex++,
                "2.3 จำนวนใบแจ้งซ่อมที่คงค้างการแก้ไขจากเดือนก่อนๆ แต่ปิดได้ในเดือนนี้", HorizontalAlignment.Left,
                VerticalAlignment.Center);
            for (int i = 2; i < 22; i++)
            {
                ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, i, "", HorizontalAlignment.Center,
                    VerticalAlignment.Center);
            }

            sheet.AddMergedRegion(new CellRangeAddress(rowIndex, rowIndex, 1, 8));

            int sum7 = 0;
            for (int i = 1; i < (valuePeriod.iMonth_current + 1); i++)
            {
                foreach (var item in count_fix)
                {
                    if (item.Month == i)
                    {
                        ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, (i + 8),
                            item.ActualFixHo, HorizontalAlignment.Center,
                            VerticalAlignment.Center);
                        sum7 += item.ActualFixHo;
                    }
                }
            }

            ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, 21, sum7, HorizontalAlignment.Center,
                VerticalAlignment.Center);

            rowIndex++;
            colIndex = 0;

            #endregion

            #region จำนวนใบแจ้งซ่อมที่สามารถแก้ไข้แล้วเสร็จทั้งสิ้นภายในเดือนที่วัดผล (ข้อ 2.1 + 2.2 + 2.3)

            row = sheet.CreateRow(rowIndex);
            ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, colIndex++, "", HorizontalAlignment.Center,
                VerticalAlignment.Center);
            ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, colIndex++,
                "จำนวนใบแจ้งซ่อมที่สามารถแก้ไข้แล้วเสร็จทั้งสิ้นภายในเดือนที่วัดผล (ข้อ 2.1 + 2.2 + 2.3)",
                HorizontalAlignment.Left, VerticalAlignment.Center);
            for (int i = 2; i < 22; i++)
            {
                ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, i, "", HorizontalAlignment.Center,
                    VerticalAlignment.Center);
            }

            sheet.AddMergedRegion(new CellRangeAddress(rowIndex, rowIndex, 1, 8));

            int sum8 = 0;
            for (int i = 1; i < (valuePeriod.iMonth_current + 1); i++)
            {
                foreach (var item in count_fix)
                {
                    if (item.Month == i)
                    {
                        ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, (i + 8),
                            item.SumActual, HorizontalAlignment.Center,
                            VerticalAlignment.Center);
                        sum8 += item.SumActual;
                    }
                }
            }

            ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, 21, sum8, HorizontalAlignment.Center,
                VerticalAlignment.Center);

            rowIndex++;
            colIndex = 0;

            #endregion

            #region 3. จำนวนใบแจ้งซ่อมที่ปิดงานได้ในเดือนที่รายงาน แต่เกินกำหนดเวลาที่ตกลงกับลูกค้า

            row = sheet.CreateRow(rowIndex);
            ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, colIndex++, "", HorizontalAlignment.Center,
                VerticalAlignment.Center);
            ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, colIndex++,
                "3. จำนวนใบแจ้งซ่อมที่ปิดงานได้ในเดือนที่รายงาน แต่เกินกำหนดเวลาที่ตกลงกับลูกค้า",
                HorizontalAlignment.Left, VerticalAlignment.Center);
            for (int i = 2; i < 22; i++)
            {
                ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, i, "", HorizontalAlignment.Center,
                    VerticalAlignment.Center);
            }

            sheet.AddMergedRegion(new CellRangeAddress(rowIndex, rowIndex, 1, 8));

            int sum9 = 0;
            for (int i = 1; i < (valuePeriod.iMonth_current + 1); i++)
            {
                foreach (var item in count_fix)
                {
                    if (item.Month == i)
                    {
                        ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, (i + 8),
                            item.FixCloseInMonth, HorizontalAlignment.Center,
                            VerticalAlignment.Center);
                        sum9 += item.FixCloseInMonth;
                    }
                }
            }

            ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, 21, sum9, HorizontalAlignment.Center,
                VerticalAlignment.Center);

            rowIndex++;
            colIndex = 0;

            #endregion

            #region 4. จำนวนใบแจ้งซ่อมที่ไม่สามารถแก้ไขแล้วเสร็จตามกำหนดเวลาที่ตกลงกับลูกค้า (คงค้างยกไปเดือนถัดไป)

            row = sheet.CreateRow(rowIndex);
            ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, colIndex++, "", HorizontalAlignment.Center,
                VerticalAlignment.Center);
            ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, colIndex++,
                "4. จำนวนใบแจ้งซ่อมที่ไม่สามารถแก้ไขแล้วเสร็จตามกำหนดเวลาที่ตกลงกับลูกค้า (คงค้างยกไปเดือนถัดไป)",
                HorizontalAlignment.Left, VerticalAlignment.Center);
            for (int i = 2; i < 22; i++)
            {
                ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, i, "", HorizontalAlignment.Center,
                    VerticalAlignment.Center);
            }

            sheet.AddMergedRegion(new CellRangeAddress(rowIndex, rowIndex, 1, 8));

            int sum10 = 0;
            for (int i = 1; i < (valuePeriod.iMonth_current + 1); i++)
            {
                foreach (var item in count_fix)
                {
                    if (item.Month == i)
                    {
                        ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, (i + 8),
                            item.FixHo, HorizontalAlignment.Center,
                            VerticalAlignment.Center);
                        sum10 += item.FixHo;
                    }
                }
            }

            ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, 21, sum10, HorizontalAlignment.Center,
                VerticalAlignment.Center);

            rowIndex++;
            colIndex = 0;

            #endregion

            #region % ใบแจ้งซ่อมที่แก้ไขแล้วเสร็จภายในกำหนดเวลาที่ตกลงกับลูกค้า (ผลรวม ข้อ 2 / 1)

            row = sheet.CreateRow(rowIndex);
            ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, colIndex++, "", HorizontalAlignment.Center,
                VerticalAlignment.Center);
            ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, colIndex++,
                "% ใบแจ้งซ่อมที่แก้ไขแล้วเสร็จภายในกำหนดเวลาที่ตกลงกับลูกค้า (ผลรวม ข้อ 2 / 1)",
                HorizontalAlignment.Left, VerticalAlignment.Center);
            for (int i = 2; i < 22; i++)
            {
                ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, i, "", HorizontalAlignment.Center,
                    VerticalAlignment.Center);
            }

            sheet.AddMergedRegion(new CellRangeAddress(rowIndex, rowIndex, 1, 8));

            double sum11 = 0;
            for (int i = 1; i < (valuePeriod.iMonth_current + 1); i++)
            {
                foreach (var item in count_fix)
                {
                    if (item.Month == i)
                    {
                        ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, (i + 8),
                            item.Perc + "%", HorizontalAlignment.Center,
                            VerticalAlignment.Center);
                        sum11 += item.Perc;
                    }
                }
            }

            sum11 = (sum11 / valuePeriod.iMonth_current);
            sum11 = (float) Math.Round((double) sum11);

            ExcelHelper.createDataTable(ref workbook, ref sheet, ref row, 21, sum11 + "%", HorizontalAlignment.Center,
                VerticalAlignment.Center);

            rowIndex++;
            colIndex = 0;

            #endregion

            #region หมายเหตุ:

            row = sheet.CreateRow(rowIndex);
            ExcelHelper.CreateDetail(ref workbook, ref sheet, ref row, colIndex++, "หมายเหตุ:", true);
            sheet.AddMergedRegion(new CellRangeAddress(rowIndex, rowIndex, 0, 21));
            rowIndex++;
            colIndex = 0;

            #endregion

            #region กรณีไม่บรรลุวัตถุประสงค์คุณภาพตามที่กำหนด กรุณาวิเคราะห์ปัญหาและกรอกข้อมูลในด้านล่าง

            row = sheet.CreateRow(rowIndex);
            colIndex = 1;
            ExcelHelper.CreateDetail(ref workbook, ref sheet, ref row, colIndex++,
                "กรณีไม่บรรลุวัตถุประสงค์คุณภาพตามที่กำหนด กรุณาวิเคราะห์ปัญหาและกรอกข้อมูลในด้านล่าง", false);
            sheet.AddMergedRegion(new CellRangeAddress(rowIndex, rowIndex, 1, 21));
            rowIndex++;
            colIndex = 0;

            #endregion

            row = sheet.CreateRow(rowIndex);
            ExcelHelper.CreateDetail(ref workbook, ref sheet, ref row, colIndex++, "", false);
            sheet.AddMergedRegion(new CellRangeAddress(rowIndex, rowIndex, 0, 21));
            rowIndex++;
            colIndex = 0;
            row = sheet.CreateRow(rowIndex);
            ExcelHelper.CreateDetail(ref workbook, ref sheet, ref row, colIndex++, "", false);
            sheet.AddMergedRegion(new CellRangeAddress(rowIndex, rowIndex, 0, 21));
            rowIndex++;
            colIndex = 0;
            row = sheet.CreateRow(rowIndex);
            ExcelHelper.CreateDetail(ref workbook, ref sheet, ref row, colIndex++, "", false);
            sheet.AddMergedRegion(new CellRangeAddress(rowIndex, rowIndex, 0, 21));
            rowIndex++;
            colIndex = 0;

            #region การแก้ไข/ป้องกันไม่ให้เกิดซ้ำ: (ผู้บังคับบัญชามีหน้าที่ติดตามให้เกิดผลในทางปฏิบัติ)

            row = sheet.CreateRow(rowIndex);
            colIndex = 1;
            ExcelHelper.CreateDetail(ref workbook, ref sheet, ref row, colIndex++,
                "การแก้ไข/ป้องกันไม่ให้เกิดซ้ำ: (ผู้บังคับบัญชามีหน้าที่ติดตามให้เกิดผลในทางปฏิบัติ)", false);
            sheet.AddMergedRegion(new CellRangeAddress(rowIndex, rowIndex, 1, 21));
            rowIndex++;
            colIndex = 0;

            #endregion

            row = sheet.CreateRow(rowIndex);
            ExcelHelper.CreateDetail(ref workbook, ref sheet, ref row, colIndex++, "", false);
            sheet.AddMergedRegion(new CellRangeAddress(rowIndex, rowIndex, 0, 21));
            rowIndex++;
            colIndex = 0;
            row = sheet.CreateRow(rowIndex);
            ExcelHelper.CreateDetail(ref workbook, ref sheet, ref row, colIndex++, "", false);
            sheet.AddMergedRegion(new CellRangeAddress(rowIndex, rowIndex, 0, 21));
            rowIndex++;
            colIndex = 0;
            row = sheet.CreateRow(rowIndex);
            ExcelHelper.CreateDetail(ref workbook, ref sheet, ref row, colIndex++, "", false);
            sheet.AddMergedRegion(new CellRangeAddress(rowIndex, rowIndex, 0, 21));
            rowIndex++;
            colIndex = 0;

            #region ผู้จัดทำ ผู้บริหารสายงาน

            row = sheet.CreateRow(rowIndex);
            colIndex = 2;
            ExcelHelper.CreateSign(ref workbook, ref sheet, ref row, colIndex++,
                (object) "ผู้จัดทำ ..................................................................................",
                true);
            for (int i = 3; i < 10; i++)
            {
                ExcelHelper.CreateSign(ref workbook, ref sheet, ref row, i, (object) "", false);
            }

            sheet.AddMergedRegion(new CellRangeAddress(rowIndex, rowIndex, 2, 9));
            colIndex = 12;
            ExcelHelper.CreateSign(ref workbook, ref sheet, ref row, colIndex++,
                (object) "ผู้บริหารสายงาน...............................................", true);
            for (int i = 13; i < 20; i++)
            {
                ExcelHelper.CreateSign(ref workbook, ref sheet, ref row, i, (object) "", false);
            }

            sheet.AddMergedRegion(new CellRangeAddress(rowIndex, rowIndex, 12, 19));
            rowIndex++;
            colIndex = 0;

            #endregion

            #region (...)

            row = sheet.CreateRow(rowIndex);
            colIndex = 2;
            ExcelHelper.CreateSign(ref workbook, ref sheet, ref row, colIndex++,
                (object) "(.......................................................................................)",
                true);
            for (int i = 3; i < 10; i++)
            {
                ExcelHelper.CreateSign(ref workbook, ref sheet, ref row, i, (object) "", false);
            }

            sheet.AddMergedRegion(new CellRangeAddress(rowIndex, rowIndex, 2, 9));
            colIndex = 12;
            ExcelHelper.CreateSign(ref workbook, ref sheet, ref row, colIndex++,
                (object) "(.......................................................................................)",
                true);
            for (int i = 13; i < 20; i++)
            {
                ExcelHelper.CreateSign(ref workbook, ref sheet, ref row, i, (object) "", false);
            }

            sheet.AddMergedRegion(new CellRangeAddress(rowIndex, rowIndex, 12, 19));
            rowIndex++;
            colIndex = 0;

            #endregion

            #region วันที่...

            row = sheet.CreateRow(rowIndex);
            colIndex = 2;
            ExcelHelper.CreateSign(ref workbook, ref sheet, ref row, colIndex++,
                (object) "(วันที่.....................................)", true);
            for (int i = 3; i < 10; i++)
            {
                ExcelHelper.CreateSign(ref workbook, ref sheet, ref row, i, (object) "", false);
            }

            sheet.AddMergedRegion(new CellRangeAddress(rowIndex, rowIndex, 2, 9));
            colIndex = 12;
            ExcelHelper.CreateSign(ref workbook, ref sheet, ref row, colIndex++,
                (object) "(วันที่.....................................)", true);
            for (int i = 13; i < 20; i++)
            {
                ExcelHelper.CreateSign(ref workbook, ref sheet, ref row, i, (object) "", false);
            }

            sheet.AddMergedRegion(new CellRangeAddress(rowIndex, rowIndex, 12, 19));
            rowIndex++;
            colIndex = 0;

            #endregion

            row = sheet.CreateRow(rowIndex);
            ExcelHelper.CreateDetail(ref workbook, ref sheet, ref row, colIndex++, "", false);
            sheet.AddMergedRegion(new CellRangeAddress(rowIndex, rowIndex, 0, 21));
            rowIndex++;
            colIndex = 0;

            #region เรียน รองกรรมการผู้จัดการ และ QMR ความเห็น รองกรรมการผู้จัดการ ความเห็น QMR

            row = sheet.CreateRow(rowIndex);
            ExcelHelper.CreateDetail(ref workbook, ref sheet, ref row, colIndex++, "เรียน รองกรรมการผู้จัดการ และ QMR",
                false);
            for (int i = 1; i < 7; i++)
            {
                ExcelHelper.CreateSign(ref workbook, ref sheet, ref row, i, (object) "", false);
            }

            sheet.AddMergedRegion(new CellRangeAddress(rowIndex, rowIndex, 0, 6));
            colIndex = 8;
            ExcelHelper.CreateComment(ref workbook, ref sheet, ref row, colIndex++, "ความเห็น รองกรรมการผู้จัดการ");
            for (int i = 9; i < 12; i++)
            {
                ExcelHelper.CreateSign(ref workbook, ref sheet, ref row, i, (object) "", false);
            }

            sheet.AddMergedRegion(new CellRangeAddress(rowIndex, rowIndex, 8, 11));
            colIndex = 15;
            ExcelHelper.CreateComment(ref workbook, ref sheet, ref row, colIndex++, "ความเห็น QMR");
            for (int i = 16; i < 22; i++)
            {
                ExcelHelper.CreateSign(ref workbook, ref sheet, ref row, i, (object) "", false);
            }

            sheet.AddMergedRegion(new CellRangeAddress(rowIndex, rowIndex, 15, 21));
            rowIndex++;
            colIndex = 0;

            #endregion

            #region ทราบ อื่น ๆ

            row = sheet.CreateRow(rowIndex);
            colIndex = 8;
            ExcelHelper.CreateComment(ref workbook, ref sheet, ref row, colIndex++, symbol);
            ExcelHelper.CreateDetail(ref workbook, ref sheet, ref row, colIndex++, "ทราบ", false);
            colIndex = 11;
            ExcelHelper.CreateComment(ref workbook, ref sheet, ref row, colIndex++, symbol);
            ExcelHelper.CreateDetail(ref workbook, ref sheet, ref row, colIndex++, "อื่น ๆ", false);
            colIndex = 15;
            ExcelHelper.CreateComment(ref workbook, ref sheet, ref row, colIndex++, "");
            ExcelHelper.CreateDetail(ref workbook, ref sheet, ref row, colIndex++, symbol, false);
            ExcelHelper.CreateDetail(ref workbook, ref sheet, ref row, colIndex++, "ทราบ", false);
            colIndex++;
            ExcelHelper.CreateDetail(ref workbook, ref sheet, ref row, colIndex++, symbol, false);
            ExcelHelper.CreateDetail(ref workbook, ref sheet, ref row, colIndex++, "ทราบ", false);
            rowIndex++;
            colIndex = 0;

            #endregion

            row = sheet.CreateRow(rowIndex);
            colIndex = 8;
            ExcelHelper.CreateComment(ref workbook, ref sheet, ref row, colIndex++, "");
            colIndex = 15;
            ExcelHelper.CreateComment(ref workbook, ref sheet, ref row, colIndex++, "");
            rowIndex++;
            colIndex = 0;
            row = sheet.CreateRow(rowIndex);
            colIndex = 8;
            ExcelHelper.CreateComment(ref workbook, ref sheet, ref row, colIndex++, "");
            colIndex = 15;
            ExcelHelper.CreateComment(ref workbook, ref sheet, ref row, colIndex++, "");
            rowIndex++;
            colIndex = 0;
            row = sheet.CreateRow(rowIndex);
            colIndex = 8;
            ExcelHelper.CreateComment(ref workbook, ref sheet, ref row, colIndex++, "");
            colIndex = 15;
            ExcelHelper.CreateComment(ref workbook, ref sheet, ref row, colIndex++, "");
            rowIndex++;
            colIndex = 0;
            row = sheet.CreateRow(rowIndex);
            colIndex = 8;
            ExcelHelper.CreateComment(ref workbook, ref sheet, ref row, colIndex++, "");
            colIndex = 15;
            ExcelHelper.CreateComment(ref workbook, ref sheet, ref row, colIndex++, "");
            rowIndex++;
            colIndex = 0;
            row = sheet.CreateRow(rowIndex);
            colIndex = 8;
            ExcelHelper.CreateComment(ref workbook, ref sheet, ref row, colIndex++, "");
            colIndex = 15;
            ExcelHelper.CreateComment(ref workbook, ref sheet, ref row, colIndex++, "");
            rowIndex++;
            colIndex = 0;

            #region หน่วยงานระบบคุณภาพ  / วันที่ ลงนาม...

            row = sheet.CreateRow(rowIndex);
            ExcelHelper.CreateDetail(ref workbook, ref sheet, ref row, colIndex++,
                "หน่วยงานระบบคุณภาพ.................................... / วันที่ ..................", false);
            colIndex = 8;
            ExcelHelper.CreateComment(ref workbook, ref sheet, ref row, colIndex++,
                "ลงนาม............................................................................");
            colIndex = 15;
            ExcelHelper.CreateComment(ref workbook, ref sheet, ref row, colIndex++,
                "ลงนาม.......................................................");
            rowIndex++;
            colIndex = 0;

            #endregion

            #region (นางวารุณี ลภิธนานุวัฒน์) (นายกิตติพงษ์ ศิริลักษณ์ตระกูล)

            row = sheet.CreateRow(rowIndex);
            colIndex = 8;
            ExcelHelper.CreateComment(ref workbook, ref sheet, ref row, colIndex++, "(นางวารุณี ลภิธนานุวัฒน์)");
            colIndex = 15;
            ExcelHelper.CreateComment(ref workbook, ref sheet, ref row, colIndex++, "(นายกิตติพงษ์ ศิริลักษณ์ตระกูล)");
            rowIndex++;
            colIndex = 0;

            #endregion

            #region ผู้อำนวยการฝ่ายกำกับดูแลกิจการฯ./ วันที่ .รองกรรมการผู้จัดการ รองกรรมการผู้จัดการ สายงานก่อสร้างอาคารสูง / QMR

            row = sheet.CreateRow(rowIndex);
            ExcelHelper.CreateDetail(ref workbook, ref sheet, ref row, colIndex++,
                "ผู้อำนวยการฝ่ายกำกับดูแลกิจการฯ.................... / วันที่ .................", false);
            colIndex = 8;
            ExcelHelper.CreateComment(ref workbook, ref sheet, ref row, colIndex++, "รองกรรมการผู้จัดการ");
            colIndex = 15;
            ExcelHelper.CreateComment(ref workbook, ref sheet, ref row, colIndex++,
                "รองกรรมการผู้จัดการสายงานก่อสร้างอาคารสูง / QMR");
            rowIndex++;
            colIndex = 0;

            #endregion

            #region วันที่...

            row = sheet.CreateRow(rowIndex);
            colIndex = 8;
            ExcelHelper.CreateComment(ref workbook, ref sheet, ref row, colIndex++,
                "วันที่.......................................................");
            colIndex = 15;
            ExcelHelper.CreateComment(ref workbook, ref sheet, ref row, colIndex++,
                "วันที่.......................................................");
            rowIndex++;
            colIndex = 0;

            #endregion

            row = sheet.CreateRow(rowIndex);
            rowIndex++;
            colIndex = 0;
        }

        public static List<CountKpiHoPeriod> GetCountKpiHoPeriod(PrmGetRpt prm)
        {
            List<CountKpiHoPeriod> result = new List<CountKpiHoPeriod>();

            var valuePeriod = BindingData.setupPeriod(prm);

            ReportDao rptDao = new ReportDao();
            var data = rptDao.GetKpiHO(prm);

            for (int i = 1; i < (valuePeriod.iMonth_current + 1); i++)
            {
                #region จำนวนแปลงที่ครบกำหนดโอนกรรมสิทธิ์ทั้งสิ้นในเดือน

                var plan = data.Where(x => x.month == i && x.year == valuePeriod.iYear_current).ToList();
                var count_plan = plan.Count();

                #endregion

                #region 2.1 ลูกค้ารับมอบได้ก่อนกำหนดโอนฯ อย่างน้อย 3 วัน (เฉพาะแปลงที่ครบกำหนดโอนฯ ในเดือนนี้)

                ///รับมอบก่อนกำหนดโอน < 3วัน ลูกค้ารับมอบ
                var act_3 = plan.Where(x =>
                        (x.date_diff.HasValue && x.date_diff >= 3 && x.handover_type == "CUSTOMER") &&
                        x.cancel == false)
                    .Count();

                #endregion

                #region 2.2 จำนวนแปลงที่ครบกำหนดโอนฯ เดือนนี้ อยู่ระหว่างการยกเลิก (ได้รับการรับมอบแทนโดยฝ่ายขาย หรือ ผู้มีอำนาจ)

                ///รับมอบก่อนกำหนดโอน < 3วัน *อยู่ระหว่าตั้งเรื่องยกเลิก
                //var act_incancel = plan.Where(x => x.cancel == true).Count();
                //&& x.date_diff.HasValue && x.date_diff >= 3
                var act_incancel = plan.Where(x => x.cancel == true)
                    .Count();

                #endregion

                #region 2.3 แปลงที่เก็บงานแล้วเสร็จ และมีการรับมอบแทนโดยฝ่ายขาย (เฉพาะที่ครบกำหนดโอนเดือนนี้)

                ///รับมอบก่อนกำหนดโอน < 3วัน ฝ่ายขายรับมอบแทน
                var act_supalai = plan.Where(x =>
                        x.date_diff.HasValue && x.date_diff >= 3 && x.handover_type == "SUPALAI" && x.cancel == false)
                    .Count();

                #endregion

                #region sum_act

                var sum_act = act_3 + act_incancel + act_supalai;

                #endregion

                #region 3. จำนวนแปลงที่ส่งมอบได้แต่เกินระยะเวลาที่กำหนด (ทั้งรับมอบโดยลูกค้าและรับมอบแทนโดยฝ่ายขาย)

                //var xx = plan.Where(x => x.date_diff < 3 //&& x.handover_date > x.contract_transfer_date
                //                        && (x.handover_date.Value.Month == i && x.handover_date.Value.Year == year_current)).ToList();
                var ho_over = plan.Where(x =>
                    x.date_diff.HasValue && x.date_diff < 3
                                         //&& x.handover_date > x.contract_transfer_date
                                         && (x.handover_date.Value.Month == i && x.handover_date.Value.Year ==
                                             valuePeriod.iYear_current)).Count();

                #endregion

                #region 4. ยอดรับมอบคงค้างในเดือน (ข้อ1 - ผลรวมข้อ 2 - ข้อ 3) (ยกไป ข้อ 1 ของ OBJ เรื่องที่ 2 เดือนถัดไป)

                ///(1-(2+3))
                var ho_null =
                    (count_plan - (sum_act + ho_over)); //plan.Where(x => !(x.handover_date.HasValue)).Count();

                #endregion

                #region % ที่ส่งมอบได้ตามเป้าหมายของเดือน (ผลรวมข้อ 2 / ข้อ1)

                float perc = 0;
                if (count_plan != 0) //md.Plan
                {
                    perc = ((float) (sum_act) / (count_plan)) * 100; // (md.SUM_HandOver / md.Plan) * 100;
                    //ii8 = (float)(Math.Round((double)ii8, 2));
                    perc = (float) Math.Round((double) perc);

                    //md.Target_PERCENT = ii8;
                }
                else
                {
                    //md.Target_PERCENT = 100;
                    perc = 100;
                }

                #endregion

                result.Add(new CountKpiHoPeriod
                {
                    PlanHo = count_plan,
                    ActualHoIn3 = act_3,
                    ActualHoInCancel = act_incancel,
                    ActualHoTypeSupalai = act_supalai,
                    SumAct = sum_act,
                    ActualHoOver = ho_over,
                    ActualHoNull = ho_null,
                    Perc = perc,
                    Month = i,
                    Year = valuePeriod.iYear_current
                });
            }

            return result;
        }

        public static List<CountKpiHoLate> GetCountKpiHoLate(PrmGetRpt prm)
        {
            List<CountKpiHoLate> result = new List<CountKpiHoLate>();

            var valuePeriod = BindingData.setupPeriod(prm);

            ReportDao rptDao = new ReportDao();
            var data = rptDao.GetKpiHO(prm);

            data = data.Where(x => x.month == valuePeriod.iMonth_current).ToList();

            var iaq = data.Select(x => new {x.proj_name, x.proj_code}).Distinct().ToList();

            foreach (var item in iaq)
            {
                string Proj_Name_ = string.Empty;
                string Proj_COUNT_ = string.Empty;
                string HandOver_COUNT_ = string.Empty;
                string HandOver_Code_ = string.Empty;
                string HandOverNULL_COUNT_ = string.Empty;
                string HandOverNULL_Code_ = string.Empty;

                #region ปัจจุบันส่งมอบแล้ว

                var ii_LESS = data.Where(x => (x.date_diff.HasValue && x.date_diff < 3
                                                                    //x.handover_date.HasValue && x.handover_date > x.contract_transfer_date
                                                                    && x.handover_date.Value.Month ==
                                                                    valuePeriod.iMonth_current &&
                                                                    x.handover_date.Value.Year ==
                                                                    valuePeriod
                                                                        .iYear_current) //(x.date_diff.HasValue && x.date_diff >= 3)                                            
                                              && x.proj_code == item.proj_code)
                    .Select(x => x.io_code).ToList();

                if (ii_LESS.Count > 0)
                {
                    HandOver_COUNT_ = ii_LESS.Count.ToString();
                }

                foreach (var i2 in ii_LESS)
                {
                    string[] io_code_ = i2.Split('-');

                    if (HandOver_Code_ != "")
                    {
                        HandOver_Code_ += "," + io_code_[1].ToString(); //i2.ToString();
                    }
                    else
                    {
                        HandOver_Code_ += io_code_[1].ToString(); //i2.ToString();
                    }
                }

                #endregion

                #region คงค้างส่งมอบ

                var ii_NULL = data.Where(x => //(!x.date_diff.HasValue || (x.date_diff.HasValue && x.date_diff < 3))
                        //&& 
                        (!x.handover_date.HasValue || (x.handover_date.HasValue &&
                                                       x.handover_date.Value.Month > valuePeriod.iMonth_current &&
                                                       x.handover_date.Value.Year == valuePeriod.iYear_current))
                        && x.proj_code == item.proj_code)
                    .Select(x => x.io_code).ToList();

                //(from list in lt_
                //         //where list.date_diff == null
                //         //&& list.project_code == item.project_code

                //     where (!list.date_diff.HasValue || (list.date_diff.HasValue && Convert.ToInt32(list.date_diff) < 10))//== null//!= null
                // && list.month == Month//Month
                // && (!list.HandOverDate.HasValue || (list.HandOverDate.HasValue && list.HandOverDate.Value.Month > Month))
                // && list.project_code == item.project_code
                //     select list.io_code).ToList();

                if (ii_NULL.Count > 0)
                {
                    HandOverNULL_COUNT_ = ii_NULL.Count.ToString();
                }

                foreach (var i2 in ii_NULL)
                {
                    string[] io_code_ = i2.Split('-');

                    if (HandOverNULL_Code_ != "")
                    {
                        HandOverNULL_Code_ += "," + io_code_[1].ToString(); //i2.ToString();
                    }
                    else
                    {
                        HandOverNULL_Code_ += io_code_[1].ToString(); //i2.ToString();
                    }
                }

                #endregion

                int iProj_COUNT_ = 0;
                if (ii_LESS.Count > 0)
                {
                    iProj_COUNT_ += ii_LESS.Count;
                }

                if (ii_NULL.Count > 0)
                {
                    iProj_COUNT_ += ii_NULL.Count;
                }

                if (iProj_COUNT_ != 0)
                {
                    result.Add(new CountKpiHoLate
                    {
                        ProjName = item.proj_name, //Proj_Name_,
                        ProjCount = iProj_COUNT_.ToString(),
                        HoCount = HandOver_COUNT_,
                        HoNullProjCode = HandOverNULL_Code_,
                        HoNullCount = HandOverNULL_COUNT_,
                        HoProjCode = HandOver_Code_,
                    });
                }
            }

            return result;
        }

        public static List<CountKpiHoNullPeriodLowRise> GetCountKpiHoNullPeriod(PrmGetRpt prm)
        {
            List<CountKpiHoNullPeriodLowRise> result = new List<CountKpiHoNullPeriodLowRise>();

            var valuePeriod = BindingData.setupPeriod(prm);

            ReportDao rptDao = new ReportDao();
            var data = rptDao.GetKpiHO(prm);

            for (int i = 1; i < (valuePeriod.iMonth_current + 1); i++)
            {
                int pre_month = 0;
                if (i == 1)
                {
                    // pre_month = i;
                    result.Add(new CountKpiHoNullPeriodLowRise
                    {
                        PlanHoNullPreMonth = 0,
                        ActHo = 0,
                        ActHoNull = 0,
                        Perc = 0,
                        MoniterHo = 0,
                        MoniterHoNull = 0,
                        HoNull = 0,
                        Month = i,
                        Year = valuePeriod.iYear_current
                    });
                }

                else
                {
                    pre_month = i - 1;

                    var plan_ho = data.Where(x => x.month == pre_month && x.year == valuePeriod.iYear_current)
                        .ToList(); //pre_month

                    ///รับมอบก่อนกำหนดโอน < 3วัน ลูกค้ารับมอบ
                    var act_3 = plan_ho.Where(x =>
                            (x.date_diff.HasValue && x.date_diff >= 3 && x.handover_type == "CUSTOMER") &&
                            x.cancel == false)
                        .ToList();

                    ///รับมอบก่อนกำหนดโอน < 3วัน *อยู่ระหว่าตั้งเรื่องยกเลิก
                    //var act_incancel = plan.Where(x => x.cancel == true).Count();
                    var act_incancel = plan_ho.Where(x => x.cancel == true && x.date_diff.HasValue && x.date_diff >= 3)
                        .ToList();

                    ///รับมอบก่อนกำหนดโอน < 3วัน ฝ่ายขายรับมอบแทน
                    var act_supalai = plan_ho.Where(x =>
                            x.date_diff.HasValue && x.date_diff >= 3 && x.handover_type == "SUPALAI" &&
                            x.cancel == false)
                        .ToList();

                    var sum_act = act_3.Count() + act_incancel.Count() + act_supalai.Count();

                    // 3. จำนวนแปลงที่ส่งมอบได้แต่เกินระยะเวลาที่กำหนด (ทั้งรับมอบโดยลูกค้าและรับมอบแทนโดยฝ่ายขาย) 
                    var ho_over = plan_ho.Where(x =>
                        x.date_diff.HasValue && x.date_diff < 3 //&& x.handover_date > x.contract_transfer_date
                                             && (x.handover_date.Value.Month == pre_month &&
                                                 x.handover_date.Value.Year ==
                                                 valuePeriod.iYear_current)).ToList();

                    #region 1. จำนวนแปลงที่คงค้างการส่งมอบจากเดือนก่อนๆ (ยกมาจากข้อ 4 ของ Obj เรื่องที่ 1 ของเดือนก่อน)

                    // 4. ยอดรับมอบคงค้างในเดือน (ข้อ1 - ผลรวมข้อ 2 - ข้อ 3) (ยกไป ข้อ 1 ของ OBJ เรื่องที่ 2 เดือนถัดไป)
                    // ///(1-(2+3)) 
                    var count_ho_null_prev_month = (plan_ho.Count() - (sum_act + ho_over.Count()));

                    var union_act_ho = act_3.Union(act_incancel).Union(act_supalai).Union(ho_over);

                    var ho_null_prev_month = plan_ho.Except(union_act_ho);


                    // var ho_null_pre_month = plan_ho.Where(x =>
                    //         (x.handover_date.HasValue && (x.handover_date.Value.Month > pre_month &&
                    //                                       x.handover_date.Value.Year == valuePeriod.iYear_current)))
                    //     .ToList();

                    #endregion

                    #region 2. จำนวนแปลงคงค้าง ที่สามารถส่งมอบได้ในเดือน

                    var act_ho = ho_null_prev_month.Where(x =>
                            (x.handover_date.HasValue && (x.handover_date.Value.Month == i &&
                                                          x.handover_date.Value.Year == valuePeriod.iYear_current)))
                        .Count();

                    #endregion

                    #region 3. จำนวนแปลงคงค้างที่ยังไม่สามารถส่งมอบได้ (ข้อ 1 - 2)

                    var act_ho_null = ho_null_prev_month.Where(x =>
                            (!x.handover_date.HasValue) || ((x.handover_date.Value.Month > i) &&
                                                            x.handover_date.Value.Year == valuePeriod.iYear_current))
                        .Count();

                    #endregion

                    #region % ที่ส่งมอบบ้านที่คงค้างได้ในเดือน (ข้อ 2 / ข้อ1)

                    float perc = 0;
                    if (ho_null_prev_month.Count() > 0) //md.Plan
                    {
                        perc = ((float) (act_ho) / (ho_null_prev_month.Count())) *
                               100; // (md.SUM_HandOver / md.Plan) * 100;
                        //ii8 = (float)(Math.Round((double)ii8, 2));
                        perc = (float) Math.Round((double) perc);
                    }
                    else
                    {
                        perc = 100;
                    }

                    #endregion


                    //var ho_premonth = data.Where(x => x.date_diff < 3 && (x.handover_date.Value.Month == pre_month && x.handover_date.Value.Year == iYear_current)).ToList();
                    var ho_null_accumulate = data.Where(x => (!x.handover_date.HasValue)
                                                             || (x.handover_date.Value.Month >
                                                                 x.contract_transfer_date.Value.Month
                                                                 && x.handover_date.Value.Year ==
                                                                 valuePeriod.iYear_current)).ToList();
                    //4. จำนวนแปลงที่ต้องส่งมอบบ้านคงค้างสะสม ณ เดือนนี้ (ข้อ 3 เดือนปัจจุบัน + ข้อ 6 ของเดือนก่อน)
                    var moniter_ho_null = 0;

                    //5. จำนวนแปลงที่คงค้าง จากเดือนก่อนๆ แต่ส่งมอบได้ในเดือนนี้ (คงค้างทุกเดือน รับมอบได้จากข้อ 6 เดือนก่อนหน้า)
                    var moniter_ho = 0;


                    #region

                    //var total_ho_null = act_ho_null.Count() + ho_premonth.Count();


                    if (i == 2)
                    {
                        moniter_ho_null = act_ho_null;
                    }
                    else
                    {
                        moniter_ho_null = act_ho_null +
                                          (ho_null_accumulate.Where(x => x.month == pre_month)
                                              .Count()); //data.Where(x => (!x.handover_date.HasValue) || x.handover_date.Value.Month == i).ToList();

                        moniter_ho = ho_null_accumulate
                            .Where(x => x.handover_date.HasValue && x.handover_date.Value.Month == i)
                            .Count(); //moniter_ho_null.Where(x => x.handover_date.HasValue && x.handover_date.Value.Month == i).Count();
                    }

                    #endregion

                    #region 5. จำนวนแปลงที่คงค้าง จากเดือนก่อนๆ แต่ส่งมอบได้ในเดือนนี้ (คงค้างทุกเดือน รับมอบได้จากข้อ 6 เดือนก่อนหน้า)

                    #endregion

                    #region 6. คงเหลือจำนวนแปลงคงค้างสะสมที่ไม่สามารถส่งมอบบ้านได้ (ข้อ 4 - 5)

                    var ho_null = moniter_ho_null - moniter_ho;
                    if (ho_null < 0)
                    {
                        ho_null = 0;
                    }

                    #endregion

                    result.Add(new CountKpiHoNullPeriodLowRise
                    {
                        PlanHoNullPreMonth = ho_null_prev_month.Count(),
                        ActHo = act_ho,
                        ActHoNull = act_ho_null,
                        Perc = perc,
                        MoniterHo = moniter_ho,
                        MoniterHoNull = moniter_ho_null,
                        HoNull = ho_null,
                        Month = i,
                        Year = valuePeriod.iYear_current
                    });
                }
            }

            return result;
        }

        public static List<CountKpiHoNullLateLowRise> GetCountKpiHoNullLateLowRise(PrmGetRpt prm)
        {
            var result = new List<CountKpiHoNullLateLowRise>();

            var valuePeriod = BindingData.setupPeriod(prm);

            ReportDao rptDao = new ReportDao();
            var data = rptDao.GetKpiHO(prm);

            for (int i = 1; i < (valuePeriod.iMonth_current + 1); i++)
            {
                int pre_month = 0;
                if (i == 1)
                {
                    pre_month = i;
                }
                else
                {
                    pre_month = i - 1;
                }

                var act_ho_null = data.Where(x =>
                        x.month == pre_month && x.year == valuePeriod.iYear_current && (!x.handover_date.HasValue))
                    .ToList();
                var iaq = act_ho_null.Select(x => new {x.proj_name, x.proj_code}).Distinct().ToList();

                foreach (var item in iaq)
                {
                    string Proj_Name_ = string.Empty;
                    string Proj_COUNT_ = string.Empty;
                    //string HandOver_COUNT_ = string.Empty;
                    //string HandOver_Code_ = string.Empty;
                    string HandOverNULL_COUNT_ = string.Empty;
                    string HandOverNULL_Code_ = string.Empty;

                    #region คงค้างส่งมอบ

                    var ii_NULL = data.Where(x => //(!x.date_diff.HasValue || (x.date_diff.HasValue && x.date_diff < 3))
                            //&& 
                            (!x.handover_date.HasValue ||
                             (x.handover_date.HasValue && x.handover_date.Value.Month > pre_month))
                            && x.proj_code == item.proj_code)
                        .Select(x => x.io_code).ToList();

                    if (ii_NULL.Count > 0)
                    {
                        HandOverNULL_COUNT_ = ii_NULL.Count.ToString();
                    }

                    foreach (var i2 in ii_NULL)
                    {
                        string[] io_code_ = i2.Split('-');

                        if (HandOverNULL_Code_ != "")
                        {
                            HandOverNULL_Code_ += "," + io_code_[1].ToString(); //i2.ToString();
                        }
                        else
                        {
                            HandOverNULL_Code_ += io_code_[1].ToString(); //i2.ToString();
                        }
                    }

                    #endregion

                    int iProj_COUNT_ = 0;
                    if (ii_NULL.Count > 0)
                    {
                        iProj_COUNT_ += ii_NULL.Count;
                    }

                    if (iProj_COUNT_ != 0)
                    {
                        result.Add(new CountKpiHoNullLateLowRise
                        {
                            ProjName = item.proj_name, //Proj_Name_,
                            ProjCount = iProj_COUNT_.ToString(),
                            //ho_count = HandOver_COUNT_,
                            HoNullProjCode = HandOverNULL_Code_,
                            HoNullCount = HandOverNULL_COUNT_,
                            //ho_projcode = HandOver_Code_,
                        });
                    }
                }
            }

            return result;
        }

        public static List<CountKpiFixLowRise> GetCountKpiFixLowRise(PrmGetRpt prm)
        {
            List<CountKpiFixLowRise> result = new List<CountKpiFixLowRise>();

            var valuePeriod = BindingData.setupPeriod(prm);

            ReportDao rptDao = new ReportDao();
            var data = rptDao.GetKpiFix(prm);

            for (int i = 1; i < (valuePeriod.iMonth_current + 1); i++)
            {
                #region 1.1 จำนวนใบแจ้งซ่อมที่กำหนดเสร็จในเดือนที่รายงาน(ไม่นับรวมใบแจ้งซ่อมที่ปิดงานและรายงานผลแล้วเมื่อเดือนก่อนๆ)

                //plan_fix_duedate_inmounth 
                var plan_fix_duedate_inmounth_ = data.Where(x => x.years == valuePeriod.iYear_current && x.months == i
                    && (
                        //pstpone null, due_date null
                        ((!x.postpone_due_date.HasValue && !x.due_date.HasValue) && x.months == i)
                        ||
                        //postpone
                        ((x.postpone_due_date.HasValue && x.due_date.HasValue) && x.postpone_due_date.Value.Month == i)
                        ||
                        //duedate
                        ((!x.postpone_due_date.HasValue && x.due_date.HasValue) && x.due_date.Value.Month == i)
                    )
                ).ToList();

                //var plan_fix_ho_ = data.Where(x => x.months < i && x.years == iYear_current && !(x.nj_closed_date.HasValue)).ToList();

                #endregion

                #region 1.2 จำนวนใบแจ้งซ่อมที่มีกำหนดแล้วเสร็จเดือนต่อๆ ไป แต่สามารถปิดงานได้ก่อนในเดือนที่รายงาน

                //plan_fix_close_inmonth
                var plan_fix_close_inmonth_ = data.Where(x => x.years == valuePeriod.iYear_current
                                                              && (
                                                                  //postpone
                                                                  ((x.postpone_due_date.HasValue &&
                                                                    x.due_date.HasValue) &&
                                                                   x.postpone_due_date.Value.Month > i)
                                                                  ||
                                                                  //duedate
                                                                  ((!x.postpone_due_date.HasValue &&
                                                                    x.due_date.HasValue) && x.due_date.Value.Month > i)
                                                              )
                                                              && x.nj_closed_date.Value.Month == i
                ).ToList();

                #endregion

                #region 1.3 จำนวนใบแจ้งซ่อมที่คงค้างการแก้ไขจากเดือนก่อน(ยกมาจาก "ข้อ 4" ของเดือนก่อนๆ)

                //plan_fix_ho
                var plan_fix_ho_ = data.Where(x => x.years == valuePeriod.iYear_current
                                                   && (
                                                       //postpone
                                                       ((x.postpone_due_date.HasValue && x.due_date.HasValue) &&
                                                        x.postpone_due_date.Value.Month < i)
                                                       ||
                                                       //duedate
                                                       ((!x.postpone_due_date.HasValue && x.due_date.HasValue) &&
                                                        x.due_date.Value.Month < i)
                                                   )
                                                   && x.nj_closed_date.Value.Month >= i
                ).ToList();

                #endregion

                #region จำนวนใบแจ้งซ่อมที่ต้องแก้ไข้แล้วเสร็จทั้งสิ้นภายในเดือนที่วัดผล(ข้อ 1.1 + ข้อ 1.2 + ข้อ 1.3)

                //sum_plan
                var sum_plan_ = plan_fix_duedate_inmounth_.Count() + plan_fix_close_inmonth_.Count() +
                                plan_fix_ho_.Count();

                #endregion

                #region 2.1 ใบแจ้งซ่อมที่มีกำหนดแล้วเสร็จในเดือนที่รายงานที่สามารถปิดได้

                //actual_duedate_inmounth
                var actual_duedate_inmounth_ = plan_fix_duedate_inmounth_.Where(x => x.nj_closed_date.Value.Month == i
                    && x.diff_date <= 0).ToList();

                #endregion

                #region 2.2 ใบแจ้งซ่อมที่มีกำหนดเสร็จในเดือนต่อๆ ไป แต่สามารถปิดได้ในเดือนที่รายงานผล

                //actual_close_inmonth
                var actual_close_inmonth_ = plan_fix_close_inmonth_.Count(); //

                #endregion

                #region 2.3 จำนวนใบแจ้งซ่อมที่คงค้างการแก้ไขจากเดือนก่อนๆ แต่ปิดได้ในเดือนนี้

                //actual_fix_ho
                var actual_fix_ho_ = plan_fix_ho_.Where(x => x.nj_closed_date.Value.Month == i).ToList();

                #endregion

                #region จำนวนใบแจ้งซ่อมที่สามารถแก้ไข้แล้วเสร็จทั้งสิ้นภายในเดือนที่วัดผล(ข้อ 2.1 + 2.2 + 2.3)

                //sum_actual
                var sum_actual_ = actual_duedate_inmounth_.Count() + actual_close_inmonth_ + actual_fix_ho_.Count();

                #endregion

                #region 3. จำนวนใบแจ้งซ่อมที่ปิดงานได้ในเดือนที่รายงาน แต่เกินกำหนดเวลาที่ตกลงกับลูกค้า

                //fix_close_inmonth
                var fix_close_inmonth_ = data.Where(x => x.nj_closed_date.Value.Month == i

                                                         ////pstpone null, due_date null
                                                         ////(!(x.postpone_due_date.HasValue) && !(x.due_date.HasValue) && x.months < i)
                                                         ////||

                                                         //postpone
                                                         && (((x.postpone_due_date.HasValue && x.due_date.HasValue) &&
                                                              x.postpone_due_date.Value.Month == i)
                                                             ||
                                                             //duedate
                                                             ((!x.postpone_due_date.HasValue && x.due_date.HasValue) &&
                                                              x.due_date.Value.Month == i))
                                                         && x.diff_date > 0
                ).ToList();

                #endregion

                #region 4. จำนวนใบแจ้งซ่อมที่ไม่สามารถแก้ไขแล้วเสร็จตามกำหนดเวลาที่ตกลงกับลูกค้า(คงค้างยกไปเดือนถัดไป)

                //fix_ho
                var fix_ho_ = data.Where(x => x.years == valuePeriod.iYear_current
                                              && (
                                                  //postpone
                                                  ((x.postpone_due_date.HasValue && x.due_date.HasValue) &&
                                                   x.postpone_due_date.Value.Month <= i)
                                                  ||
                                                  //duedate
                                                  ((!x.postpone_due_date.HasValue && x.due_date.HasValue) &&
                                                   x.due_date.Value.Month <= i)
                                              )
                                              && x.nj_closed_date.Value.Month > i
                    //&& !(x.nj_closed_date.HasValue)
                ).ToList();

                #endregion

                #region % ใบแจ้งซ่อมที่แก้ไขแล้วเสร็จภายในกำหนดเวลาที่ตกลงกับลูกค้า(ผลรวม ข้อ 2 / 1)

                //perc
                float perc = 0;
                if (sum_plan_ != 0) //md.Plan
                {
                    perc = ((float) (sum_actual_) / (sum_plan_)) * 100; // (md.SUM_HandOver / md.Plan) * 100;
                    //ii8 = (float)(Math.Round((double)ii8, 2));
                    perc = (float) Math.Round((double) perc);

                    //md.Target_PERCENT = ii8;
                }
                else
                {
                    //md.Target_PERCENT = 100;
                    perc = 100;
                }

                #endregion

                result.Add(new CountKpiFixLowRise()
                {
                    PlanFixDuedateInMounth = plan_fix_duedate_inmounth_.Count(),
                    PlanFixCloseInMonth = plan_fix_close_inmonth_.Count(),
                    PlanFixHo = plan_fix_ho_.Count(),
                    SumPlan = sum_plan_,
                    ActualDuedateInMounth = actual_duedate_inmounth_.Count(),
                    ActualCloseInMonth = actual_close_inmonth_,
                    ActualFixHo = actual_fix_ho_.Count(),
                    SumActual = sum_actual_,
                    FixCloseInMonth = fix_close_inmonth_.Count(),
                    FixHo = fix_ho_.Count(),
                    Perc = perc,

                    Month = i,
                    Year = valuePeriod.iYear_current
                });
            }

            return result;
        }
    }
}