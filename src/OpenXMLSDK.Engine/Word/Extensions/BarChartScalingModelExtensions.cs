using System.Collections.Generic;
using System.Diagnostics;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing.Charts;
using OpenXMLSDK.Engine.Word.ReportEngine.Models.Charts;

namespace OpenXMLSDK.Engine.Word.Extensions
{
    public static class BarChartScalingModelExtensions
    {
        /// <summary>
        /// Construct a DocumentFormat.OpenXml.Drawing.Charts.Scaling object
        /// </summary>
        /// <returns></returns>
        public static Scaling GetScaling(this BarChartScalingModel model)
        {
            if (model is null)
                return new Scaling() { Orientation = new Orientation() { Val = OrientationValues.MinMax } };

            var scalingParams = new List<OpenXmlElement>()
            {
                new Orientation()
                {
                    Val = model.Orientation.ToOxmlEnumValue()
                }
            };

            if (model.MinAxisValue.HasValue)
            {
                scalingParams.Add(new MinAxisValue()
                {
                    Val = new DoubleValue(model.MinAxisValue.Value)
                });
            }

            if (model.MaxAxisValue.HasValue)
            {
                scalingParams.Add(new MaxAxisValue()
                {
                    Val = new DoubleValue(model.MaxAxisValue.Value)
                });
            }

            return new Scaling(scalingParams);
        }

        private static OrientationValues ToOxmlEnumValue(this BarChartOrientationType orientation)
        {
            if(orientation == BarChartOrientationType.MaxMin)
            {
                return OrientationValues.MaxMin;
            }

            return OrientationValues.MinMax;
        }
    }
}
