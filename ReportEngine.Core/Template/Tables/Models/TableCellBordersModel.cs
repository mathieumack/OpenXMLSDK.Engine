namespace ReportEngine.Core.Template.Tables.Models
{
    public class TableCellBordersModel
    {
        public TableBorderModel TopBorder { get; set; }

        public TableBorderModel BottomBorder { get; set; }

        public TableBorderModel LeftBorder { get; set; }

        public TableBorderModel RightBorder { get; set; }

        public TableBorderModel TopLeftToBottomRightCellBorder { get; set; }

        public TableBorderModel TopRightToBottomLeftCellBorder { get; set; }
    }
}