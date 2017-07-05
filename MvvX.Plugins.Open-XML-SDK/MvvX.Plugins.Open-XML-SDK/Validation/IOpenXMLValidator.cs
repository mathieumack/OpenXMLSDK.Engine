using System.Collections.Generic;

namespace MvvX.Plugins.OpenXMLSDK.Validation
{
    public interface IOpenXMLValidator
    {
        /// <summary>
        /// Validate a docx document and return all finded errors
        /// </summary>
        /// <param name="filePath"></param>
        /// <returns></returns>
        IEnumerable<ValidationErrorInfo> ValidateWordDocument(string filePath);
    }
}
