using KpiISO.Data.Model;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;

namespace KpiISO.App_Helpers
{
    public static class ExcelHelper
    {
        static string fontName = "Cardia New";
        
        public static void createHeader(ref XSSFWorkbook workbook, ref ISheet sheet, ref IRow row, int icol, object cellValue, HorizontalAlignment horizontalAlignment, VerticalAlignment verticalAlignment)
        {
            row.HeightInPoints = 30;


            #region Font
            IFont font = workbook.CreateFont();
            setupFontStyle(ref font, true, 16, fontName);

            #endregion
            #region Alignment
            var cellAlignment = setupAlignment(horizontalAlignment, verticalAlignment);

            #endregion
            #region Border
            var cellBorder = new CellBorder();
            cellBorder.BorderLeft = 0; cellBorder.BorderRight = 0; cellBorder.BorderBottom = 0; cellBorder.BorderTop = 0;

            #endregion

            var cell = row.CreateCell(icol);

            cell.SetCellValue(cellValue.ToString());
            cell.CellStyle = CreateCellStyle(ref workbook, ref font, cellAlignment, cellBorder, true);

        }
        public static void createHeaderDetail(ref XSSFWorkbook workbook, ref ISheet sheet, ref IRow row, int icol, object cellValue, HorizontalAlignment horizontalAlignment, VerticalAlignment verticalAlignment)
        {
            row.HeightInPoints = 30;

            #region Font
            IFont font = workbook.CreateFont();
            setupFontStyle(ref font, true, 16, fontName);

            #endregion            
            #region Border
            var cellBorder = new CellBorder();
            cellBorder.BorderLeft = 0; cellBorder.BorderRight = 0; cellBorder.BorderBottom = 0; cellBorder.BorderTop = 0;

            #endregion
            #region Alignment
            var cellAlignment = setupAlignment(HorizontalAlignment.Left, VerticalAlignment.Center);

            #endregion

            var cell = row.CreateCell(icol);

            cell.SetCellValue(cellValue.ToString());
            cell.CellStyle = CreateCellStyle(ref workbook, ref font, cellAlignment, cellBorder, false);

        }
        public static CellAlignment setupAlignment(HorizontalAlignment horizontalAlignment, VerticalAlignment verticalAlignment)
        {
            var cellAlignment = new CellAlignment();
            cellAlignment.horizontalAlignment = horizontalAlignment;
            cellAlignment.verticalAlignment = verticalAlignment;
            return cellAlignment;
        }
        public static void createEmptyCell(ref XSSFWorkbook workbook, ref ISheet sht_handover, ref IRow row, int icol)
        {
            var cell = row.CreateCell(icol);
        }
        public static void createDataTable(ref XSSFWorkbook workbook, ref ISheet sheet, ref IRow row, int icol, object cellValue, HorizontalAlignment horizontalAlignment, VerticalAlignment verticalAlignment)
        {
            try
            {
                if (cellValue == null)
                {
                    cellValue = new object();
                }
                Type type = cellValue.GetType();

                int n;
                bool isNumeric = int.TryParse(cellValue.ToString(), out n);

                var cell = row.CreateCell(icol);
                //sheet.AutoSizeColumn(icol, isNumeric);

                if (type == typeof(object))
                {
                    cell.SetCellValue("");
                }
                else if (type == typeof(int) || type == typeof(Int16)
                     || type == typeof(Int32) || type == typeof(Int64) || isNumeric)
                {
                    cell.SetCellValue(Convert.ToInt32(cellValue));
                }
                else if (type == typeof(float) || type == typeof(double) || type == typeof(Double))
                {
                    cell.SetCellValue((Double)cellValue);
                }
                else if (type == typeof(DateTime))
                {
                    cell.SetCellValue(((DateTime)cellValue).ToString("dd/MM/yyyy"));/*HH:mm*/
                }
                else if (type == typeof(bool) || type == typeof(Boolean))
                {
                    cell.SetCellValue((bool)cellValue);
                }
                else if (type == typeof(string))
                {
                    cell.SetCellValue(cellValue.ToString());
                }
                else
                {
                    cell.SetCellValue(cellValue.ToString());
                }

                #region Font
                IFont font = workbook.CreateFont();
                setupFontStyle(ref font, true, 14, fontName);

                #endregion
                #region Alignment
                var cellAlignment = setupAlignment(horizontalAlignment, verticalAlignment);

                #endregion
                #region Border
                var cellBorder = new CellBorder();
                cellBorder.BorderLeft = BorderStyle.Thin; cellBorder.BorderRight = BorderStyle.Thin;
                cellBorder.BorderBottom = BorderStyle.Thin; cellBorder.BorderTop = BorderStyle.Thin;

                #endregion

                cell.CellStyle = CreateCellStyle(ref workbook, ref font, cellAlignment, cellBorder, false);
                cell.Row.Height = 600;

            }
            catch (Exception ex)
            {
                throw ex;
            }

        }  
        public static ICellStyle CreateCellStyle(ref XSSFWorkbook workbook, ref IFont font, CellAlignment cellAlignment, CellBorder cellBorder, bool WrapText)
        {
            var cellStyle = (XSSFCellStyle)workbook.CreateCellStyle();
            try
            {
                //IFont font = workbook.CreateFont();
                //setupFontStyle(ref font, isBold, fontHeightInPoints);

                cellStyle.SetFont(font);
                cellStyle.Alignment = cellAlignment.horizontalAlignment;
                cellStyle.VerticalAlignment = cellAlignment.verticalAlignment;

                #region Border
                cellStyle.BorderLeft = cellBorder.BorderLeft;
                cellStyle.BorderTop = cellBorder.BorderTop;
                cellStyle.BorderRight = cellBorder.BorderRight;
                cellStyle.BorderBottom = cellBorder.BorderBottom;

                #endregion 

                cellStyle.WrapText = WrapText;

                //cellStyle.FillPattern = FillPattern.SolidForeground;
            }
            catch (Exception ex)
            {
                throw ex;
            }

            return cellStyle;
        }
        public static void setupBorderStyle(ref XSSFCellStyle cellStyle, BorderStyle BorderLeft, BorderStyle BorderTop, BorderStyle BorderRight, BorderStyle BorderBottom)
        {
            cellStyle.BorderLeft = BorderLeft;
            cellStyle.BorderTop = BorderTop;
            cellStyle.BorderRight = BorderRight;
            cellStyle.BorderBottom = BorderBottom;
        }
        public static void setupFontStyle(ref IFont font, bool isBold, int fontHeightInPoints, string FontName)
        {
            font.IsBold = isBold;
            font.FontHeightInPoints = fontHeightInPoints;
            font.FontName = FontName;
        }
        public static void setupFontStyle(ref IFont font, bool isBold, int fontHeightInPoints, string FontName, FontUnderlineType fontUnderlineType)
        {
            font.IsBold = isBold;
            font.FontHeightInPoints = fontHeightInPoints;
            font.FontName = FontName;
            font.Underline = FontUnderlineType.Single;
        } 
        public static void CreateDetail(ref XSSFWorkbook workbook, ref ISheet sheet, ref IRow row, int icol, object cellValue, bool IsUnderLine)
        {
            try
            {
                if (cellValue == null)
                {
                    cellValue = new object();
                }
                Type type = cellValue.GetType();

                int n;
                bool isNumeric = int.TryParse(cellValue.ToString(), out n);

                var cell = row.CreateCell(icol);
                //sheet.AutoSizeColumn(icol, isNumeric);

                if (type == typeof(object))
                {
                    cell.SetCellValue("");
                }
                else if (type == typeof(int) || type == typeof(Int16)
                     || type == typeof(Int32) || type == typeof(Int64) || isNumeric)
                {
                    cell.SetCellValue(Convert.ToInt32(cellValue));
                }
                else if (type == typeof(float) || type == typeof(double) || type == typeof(Double))
                {
                    cell.SetCellValue((Double)cellValue);
                }
                else if (type == typeof(DateTime))
                {
                    cell.SetCellValue(((DateTime)cellValue).ToString("dd/MM/yyyy"));/*HH:mm*/
                }
                else if (type == typeof(bool) || type == typeof(Boolean))
                {
                    cell.SetCellValue((bool)cellValue);
                }
                else if (type == typeof(string))
                {
                    cell.SetCellValue(cellValue.ToString());
                }
                else
                {
                    cell.SetCellValue(cellValue.ToString());
                }

                #region Font
                IFont font = workbook.CreateFont();
                if (IsUnderLine)
                {
                    setupFontStyle(ref font, true, 14, fontName, FontUnderlineType.Single);
                }
                else
                {
                    setupFontStyle(ref font, true, 14, fontName);
                }

                #endregion
                #region Alignment
                var cellAlignment = setupAlignment(HorizontalAlignment.Left, VerticalAlignment.Center);

                #endregion
                #region Border
                var cellBorder = new CellBorder();
                cellBorder.BorderLeft = BorderStyle.None; cellBorder.BorderRight = BorderStyle.None;
                cellBorder.BorderBottom = BorderStyle.None; cellBorder.BorderTop = BorderStyle.None;

                #endregion

                cell.CellStyle = CreateCellStyle(ref workbook, ref font, cellAlignment, cellBorder, false);
                cell.Row.Height = 600;

            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        public static void CreateSign(ref XSSFWorkbook workbook, ref ISheet sheet, ref IRow row, int icol, object cellValue, bool IsUnderLine)
        {
            try
            {
                if (cellValue == null)
                {
                    cellValue = new object();
                }
                Type type = cellValue.GetType();

                int n;
                bool isNumeric = int.TryParse(cellValue.ToString(), out n);

                var cell = row.CreateCell(icol);
                //sheet.AutoSizeColumn(icol, isNumeric);

                if (type == typeof(object))
                {
                    cell.SetCellValue("");
                }
                else if (type == typeof(int) || type == typeof(Int16)
                     || type == typeof(Int32) || type == typeof(Int64) || isNumeric)
                {
                    cell.SetCellValue(Convert.ToInt32(cellValue));
                }
                else if (type == typeof(float) || type == typeof(double) || type == typeof(Double))
                {
                    cell.SetCellValue((Double)cellValue);
                }
                else if (type == typeof(DateTime))
                {
                    cell.SetCellValue(((DateTime)cellValue).ToString("dd/MM/yyyy"));/*HH:mm*/
                }
                else if (type == typeof(bool) || type == typeof(Boolean))
                {
                    cell.SetCellValue((bool)cellValue);
                }
                else if (type == typeof(string))
                {
                    cell.SetCellValue(cellValue.ToString());
                }
                else
                {
                    cell.SetCellValue(cellValue.ToString());
                }

                #region Font
                IFont font = workbook.CreateFont();
                if (IsUnderLine)
                {
                    setupFontStyle(ref font, true, 14, fontName, FontUnderlineType.Single);
                }
                else
                {
                    setupFontStyle(ref font, true, 14, fontName);
                }

                #endregion
                #region Alignment
                var cellAlignment = setupAlignment(HorizontalAlignment.Center, VerticalAlignment.Center);

                #endregion
                #region Border
                var cellBorder = new CellBorder();
                cellBorder.BorderLeft = BorderStyle.None; cellBorder.BorderRight = BorderStyle.None;
                cellBorder.BorderBottom = BorderStyle.None; cellBorder.BorderTop = BorderStyle.None;

                #endregion

                cell.CellStyle = CreateCellStyle(ref workbook, ref font, cellAlignment, cellBorder, false);
                cell.Row.Height = 600;

            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        public static void CreateComment(ref XSSFWorkbook workbook, ref ISheet sheet, ref IRow row, int icol, object cellValue)
        {
            try
            {
                if (cellValue == null)
                {
                    cellValue = new object();
                }
                Type type = cellValue.GetType();

                int n;
                bool isNumeric = int.TryParse(cellValue.ToString(), out n);

                var cell = row.CreateCell(icol);
                //sheet.AutoSizeColumn(icol, isNumeric);

                if (type == typeof(object))
                {
                    cell.SetCellValue("");
                }
                else if (type == typeof(int) || type == typeof(Int16)
                     || type == typeof(Int32) || type == typeof(Int64) || isNumeric)
                {
                    cell.SetCellValue(Convert.ToInt32(cellValue));
                }
                else if (type == typeof(float) || type == typeof(double) || type == typeof(Double))
                {
                    cell.SetCellValue((Double)cellValue);
                }
                else if (type == typeof(DateTime))
                {
                    cell.SetCellValue(((DateTime)cellValue).ToString("dd/MM/yyyy"));/*HH:mm*/
                }
                else if (type == typeof(bool) || type == typeof(Boolean))
                {
                    cell.SetCellValue((bool)cellValue);
                }
                else if (type == typeof(string))
                {
                    cell.SetCellValue(cellValue.ToString());
                }
                else
                {
                    cell.SetCellValue(cellValue.ToString());
                }

                #region Font
                IFont font = workbook.CreateFont();
                setupFontStyle(ref font, true, 14, fontName);

                #endregion
                #region Alignment
                var cellAlignment = setupAlignment(HorizontalAlignment.Left, VerticalAlignment.Center);

                #endregion
                #region Border
                var cellBorder = new CellBorder();
                cellBorder.BorderLeft = BorderStyle.DashDot; cellBorder.BorderRight = BorderStyle.None;
                cellBorder.BorderBottom = BorderStyle.None; cellBorder.BorderTop = BorderStyle.None;

                #endregion

                cell.CellStyle = CreateCellStyle(ref workbook, ref font, cellAlignment, cellBorder, false);
                cell.Row.Height = 600;

            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        public static void SetupSheet(ref ISheet sheet)
        {
            sheet.SetZoom(70);
            sheet.PrintSetup.PaperSize = (short)PaperSize.A4 + 1;
            //10.86
        }
 
    }
}