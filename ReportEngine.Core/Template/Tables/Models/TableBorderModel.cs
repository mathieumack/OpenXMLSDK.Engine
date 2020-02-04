using System.Linq;

namespace ReportEngine.Core.Template.Tables.Models
{
    public class TableBorderModel
    {
        /// <summary>
        /// Type of borders.
        /// Default : BorderValues.Single
        /// </summary>
        public BorderValues BorderValue { get; set; }

        /// <summary>
        /// Size of the order.
        /// Default : 1
        /// </summary>
        public int? Size { get; set; }

        /// <summary>
        /// Color of the border
        /// Default : Colors.Black
        /// </summary>
        public string Color { get; set; }

        public TableBorderModel()
        {
            Size = 1;
            Color = Colors.Black;
            BorderValue = BorderValues.Single;
        }
    }
}
