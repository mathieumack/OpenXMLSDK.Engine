﻿using A = DocumentFormat.OpenXml.Drawing;
using DC = DocumentFormat.OpenXml.Drawing.Charts;

namespace OpenXMLSDK.Engine.Word.ReportEngine.Models.Charts.Extensions
{
    public static class ChartExtensions
    {
        /// <summary>
        /// Add a solid fill as shape property if the color is correctly formatted (not null and hexa format)
        /// </summary>
        /// <param name="chart"></param>
        /// <param name="pieModel"></param>
        public static DC.Chart TryAddTitle(this DC.Chart chart, PieModel pieModel)
        {
            if (chart is null || pieModel is null || !pieModel.ShowTitle)
                return chart; // Nothing to do

            var titleChart = chart.AppendChild<DC.Title>(new DC.Title());
            titleChart.AppendChild(new DC.ChartText(new DC.RichText(
                new A.BodyProperties(),
                new A.ListStyle(),
                new A.Paragraph(new A.Run(new A.Text(pieModel.Title))))));
            titleChart.AppendChild(new DC.Overlay() { Val = false });

            return chart;
        }
    }
}
