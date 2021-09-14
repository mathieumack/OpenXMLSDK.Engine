using A = DocumentFormat.OpenXml.Drawing;
using dc = DocumentFormat.OpenXml.Drawing.Charts;
using OpenXMLSDK.Engine.Word.ReportEngine.Models.Charts;

namespace OpenXMLSDK.Engine.Word.ReportEngine.Models.Charts.Extensions
{
    public static class ChartExtensions
    {
        /// <summary>
        /// Add a solid fill as shape property if the color is correctly formatted (not null and hexa format)
        /// </summary>
        /// <param name="chart"></param>
        /// <param name="pieModel"></param>
        public static dc.Chart TryAddTitle(this dc.Chart chart, PieModel pieModel)
        {
            if (chart is null || pieModel is null || !pieModel.ShowTitle)
                return chart; // Nothing to do

            var titleChart = chart.AppendChild<dc.Title>(new dc.Title());
            titleChart.AppendChild(new dc.ChartText(new dc.RichText(
                new A.BodyProperties(),
                new A.ListStyle(),
                new A.Paragraph(new A.Run(new A.Text(pieModel.Title))))));
            titleChart.AppendChild(new dc.Overlay() { Val = false });

            return chart;
        }
    }
}
