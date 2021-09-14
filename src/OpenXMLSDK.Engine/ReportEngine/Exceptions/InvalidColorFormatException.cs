using System;
using System.Collections.Generic;
using System.Text;

namespace OpenXMLSDK.Engine.ReportEngine.Exceptions
{
    [Serializable]
    public class InvalidColorFormatException : Exception
    {
        public InvalidColorFormatException(string color)
            : base($"{color} is not correctly formatted. It should contains only 6 chars")
        {

        }
    }
}
