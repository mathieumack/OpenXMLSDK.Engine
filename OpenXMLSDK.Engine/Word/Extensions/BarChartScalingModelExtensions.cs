using System.Collections.Generic;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing.Charts;
using ReportEngine.Core.Template.Charts;

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
            var scalingParams = new List<OpenXmlElement>()
            {
                new Orientation()
                {
                    Val = new EnumValue<OrientationValues>((OrientationValues)(int)model.Orientation)
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
    }
}
