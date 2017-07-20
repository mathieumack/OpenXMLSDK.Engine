using System.Collections.Generic;

namespace MvvX.Plugins.OpenXMLSDK.Word.Models.Charts
{
    public class ChartSerie
    {
        /// <summary>
        /// Name of the serie
        /// </summary>
        public string Name { get; set; }

        /// <summary>
        /// Color
        /// </summary>
        public string Color { get; set; }

        /// <summary>
        /// Values
        /// </summary>
        public List<string> Values { get; set; }
    }
}
