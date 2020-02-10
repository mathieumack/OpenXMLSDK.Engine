using System.Globalization;
using iTextSharp.text;

namespace Pdf.Engine.ReportEngine.Helpers
{
    public static class FontHelper
    {
        public static BaseColor ConverPdfColorToColor(string color)
        {
            if (string.IsNullOrWhiteSpace(color))
                return default(BaseColor);

            //replace # occurences
            if (color.IndexOf('#') != -1)
                color = color.Replace("#", "");

            int r, g, b = 0;

            r = int.Parse(color.Substring(0, 2), NumberStyles.AllowHexSpecifier);
            g = int.Parse(color.Substring(2, 2), NumberStyles.AllowHexSpecifier);
            b = int.Parse(color.Substring(4, 2), NumberStyles.AllowHexSpecifier);

            return new BaseColor(r, g, b);
        }
    }
}
