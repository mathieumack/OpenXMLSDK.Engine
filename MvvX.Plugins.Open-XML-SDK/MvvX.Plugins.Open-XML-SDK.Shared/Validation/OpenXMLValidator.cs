using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using MvvX.Plugins.OpenXMLSDK.Validation;

namespace MvvX.Plugins.OpenXMLSDK.Platform.Validation
{
    public class OpenXMLValidator : IOpenXMLValidator
    {
        public IEnumerable<OpenXMLSDK.Validation.ValidationErrorInfo> ValidateWordDocument(string filePath)
        {
                using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(filePath, false))
                {
                    var validator = new OpenXmlValidator();
                    var errors = validator.Validate(wordDoc);
                    return errors.Select(e => new OpenXMLSDK.Validation.ValidationErrorInfo()
                    {
                        Id = e.Id,
                        ErrorType = (OpenXMLSDK.Validation.ValidationErrorType)(int)e.ErrorType,
                        Description = e.Description,
                        XmlPath = e.Path.XPath
                    });
            }
        }
    }
}
