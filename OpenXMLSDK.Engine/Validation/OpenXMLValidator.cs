using System.Collections.Generic;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;

namespace OpenXMLSDK.Engine.Validation
{
    public static class OpenXmlValidator
    {
        /// <summary>
        /// Validate a word document and return all detected errors
        /// </summary>
        /// <param name="filePath"></param>
        /// <returns></returns>
        public static IEnumerable<ValidationErrorInfo> ValidateWordDocument(string filePath)
        {
            using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(filePath, false))
            {
                var validator = new DocumentFormat.OpenXml.Validation.OpenXmlValidator();
                var errors = validator.Validate(wordDoc);
                return errors;
            }
        }
    }
}