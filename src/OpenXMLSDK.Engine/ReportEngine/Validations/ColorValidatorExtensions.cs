using OpenXMLSDK.Engine.ReportEngine.Exceptions;
using System.Text.RegularExpressions;

namespace OpenXMLSDK.Engine.ReportEngine.Validations
{
    public static class ColorValidatorExtensions
    {
        /// <summary>
        /// Check that <paramref name="color"/> is correctly formatted (ex : 000000)
        /// </summary>
        /// <param name="color"></param>
        /// <exception cref="InvalidColorFormatException">Invalid format (null of invalid)</exception>
        public static void CheckColorFormat(this string color)
        {
            if (string.IsNullOrWhiteSpace(color) || !Regex.IsMatch(color, "^[0-9-A-F]{6}$"))
                throw new InvalidColorFormatException(color);
        }
    }
}
