using DocumentFormat.OpenXml.Drawing;
using OpenXMLSDK.Engine.ReportEngine.Validations;
using OpenXMLSDK.Engine.Word.ReportEngine.Models.Charts;

namespace OpenXMLSDK.Engine.Word.ReportEngine.Models.Charts.Extensions
{
    public static class ShapePropertiesExtensions
    {
        /// <summary>
        /// Add a solid fill as shape property if the color is correctly formatted (not null and hexa format)
        /// </summary>
        /// <param name="shapeProperties"></param>
        /// <param name="serie"></param>
        public static ShapeProperties TryAddSolidLine(this ShapeProperties shapeProperties, PieSerie serie)
        {
            if (serie is null || string.IsNullOrWhiteSpace(serie.Color))
                return shapeProperties; // Nothing to do

            var color = serie.Color.Replace("#", "");
            color.CheckColorFormat();

            shapeProperties.AppendChild(new SolidFill() { RgbColorModelHex = new RgbColorModelHex() { Val = color } });

            return shapeProperties;
        }

        /// <summary>
        /// Add a solid fill as shape property if the color is correctly formatted (not null and hexa format)
        /// </summary>
        /// <param name="shapeProperties"></param>
        /// <param name="color"></param>
        public static ShapeProperties TryAddBorder(this ShapeProperties shapeProperties, PieSerie serie)
        {
            if (serie is null || !serie.HasBorder)
                return shapeProperties; // Nothing to do

            serie.BorderWidth = serie.BorderWidth.HasValue ? serie.BorderWidth.Value : 12700;

            serie.BorderColor = !string.IsNullOrEmpty(serie.BorderColor) ? serie.BorderColor : "000000";
            serie.BorderColor = serie.BorderColor.Replace("#", "");
            serie.BorderColor.CheckColorFormat();

            shapeProperties.AppendChild(new Outline(new SolidFill(new RgbColorModelHex() { Val = serie.BorderColor })) { Width = serie.BorderWidth.Value });

            return shapeProperties;
        }
    }
}
