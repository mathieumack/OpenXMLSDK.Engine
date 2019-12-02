using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing.Charts;
using System.Collections.Generic;

namespace OpenXMLSDK.Engine.Word.ReportEngine.Models.Charts
{
    /// <summary>
    /// Class representating the scaling of a bar chart axis
    /// </summary>
    public class BarChartScaling
    {
        /// <summary>
        /// Orientation of the axis
        /// </summary>
        public EnumValue<OrientationValues> Orientation { get; set; }

        /// <summary>
        /// Minimum axis value
        /// </summary>
        public double? MinAxisValue { get; set; }

        /// <summary>
        /// Maximum axis value
        /// </summary>
        public double? MaxAxisValue { get; set; }

        /// <summary>
        /// Constructor, set the default orientation to Min-Max
        /// </summary>
        public BarChartScaling()
        {
            Orientation = new EnumValue<OrientationValues>(OrientationValues.MinMax);
        }

        /// <summary>
        /// Construct a DocumentFormat.OpenXml.Drawing.Charts.Scaling object
        /// </summary>
        /// <returns></returns>
        public Scaling GetScaling()
        {
            var scalingParams = new List<OpenXmlElement>()
            {
                new Orientation()
                {
                    Val = Orientation
                }
            };

            if (MinAxisValue.HasValue && MaxAxisValue.HasValue)
            {
                scalingParams.Add(new MinAxisValue()
                {
                    Val = new DoubleValue(MinAxisValue.Value)
                });
                scalingParams.Add(new MaxAxisValue()
                {
                    Val = new DoubleValue(MaxAxisValue.Value)
                });
            }

            return new Scaling(scalingParams);
        }
    }
}
