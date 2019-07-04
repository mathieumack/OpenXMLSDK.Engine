using System.Collections.Generic;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;

namespace OpenXMLSDK.Engine.Validation
{
    public static class OpenXMLValidator
    {
        public static IEnumerable<ValidationErrorInfo> ValidateWordDocument(string filePath)
        {
            using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(filePath, false))
            {
                var validator = new OpenXmlValidator();
                var errors = validator.Validate(wordDoc);
                return errors;
            }
        }
    }
}