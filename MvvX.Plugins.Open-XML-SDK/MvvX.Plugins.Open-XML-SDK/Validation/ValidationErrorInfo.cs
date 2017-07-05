namespace MvvX.Plugins.OpenXMLSDK.Validation
{
    public class ValidationErrorInfo
    {
        /// <summary>
        /// Identifier of the error
        /// </summary>
        public string Id { get; set; }

        /// <summary>
        /// Type of the error
        /// </summary>
        public ValidationErrorType ErrorType { get; set; }

        /// <summary>
        /// Description of the error
        /// </summary>
        public string Description { get; set; }

        /// <summary>
        /// XmlPath to the element
        /// </summary>
        public string XmlPath { get; set; }
    }
}
