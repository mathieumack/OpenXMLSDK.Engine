using Microsoft.VisualStudio.TestTools.UnitTesting;
using OpenXMLSDK.Engine.ReportEngine.Exceptions;
using OpenXMLSDK.Engine.ReportEngine.Validations;

namespace OpenXMLSDK.UnitTest.ReportEngine.Validations
{
    [TestClass]
    public class ColorValidatorExtensions
    {
        [TestMethod]
        [ExpectedException(typeof(InvalidColorFormatException))]
        public void NullColor_Invalid()
        {
            string color = null;
            color.CheckColorFormat();
        }

        [TestMethod]
        [ExpectedException(typeof(InvalidColorFormatException))]
        public void EmptyColor_Invalid()
        {
            string color = string.Empty;
            color.CheckColorFormat();
        }

        [TestMethod]
        [ExpectedException(typeof(InvalidColorFormatException))]
        public void InvalidFormatColor_Invalid()
        {
            string color = "Hello";
            color.CheckColorFormat();
        }

        [TestMethod]
        public void ValidFormatColor_Valid()
        {
            string color = "A2F456";
            color.CheckColorFormat();
        }
    }
}
