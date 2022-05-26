using NPOI.SS.UserModel;

namespace KpiISO.Data.Model
{
    public class ExcelModel
    {
    }
    public class CellAlignment
    {
        public HorizontalAlignment horizontalAlignment { get; set; }
        public VerticalAlignment verticalAlignment { get; set; }
    }
    public class CellBorder
    {
        public BorderStyle BorderLeft { get; set; }
        public BorderStyle BorderTop { get; set; }
        public BorderStyle BorderRight { get; set; }
        public BorderStyle BorderBottom { get; set; }
    }
}