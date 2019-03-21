using System;

namespace OpenXMLSDK.Engine.Word.Charts
{
    public class ChartModelException : Exception
    {
        /// <summary>
        /// Error code (internal codes)
        /// </summary>
        public string ErrorCode { get; private set; }

        public ChartModelException(string errorCode)
            :base()
        {
            this.ErrorCode = errorCode;
        }
        public ChartModelException(string message, string errorCode)
            : base(message)
        {
            this.ErrorCode = errorCode;
        }

        public ChartModelException(string message, string errorCode, Exception innerException)
            : base(message, innerException)
        {
            this.ErrorCode = errorCode;
        }
    }
}
