﻿using OpenXMLSDK.Engine.Word;

namespace OpenXMLSDK.Engine.Platform.Word.Extensions
{
    /// <summary>
    /// extension class for the stylevalues enumeration
    /// </summary>
    public static class StyleValuesExtensions
    {
        /// <summary>
        /// convert the StyleValue enum.
        /// </summary>
        /// <param name="value"></param>
        /// <returns></returns>
        public static DocumentFormat.OpenXml.Wordprocessing.StyleValues ToOOxml(this StyleValues value)
        {
            return new DocumentFormat.OpenXml.Wordprocessing.StyleValues(value.ToString().ToLower());
        }
    }
}
