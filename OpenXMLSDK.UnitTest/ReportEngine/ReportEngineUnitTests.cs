using Microsoft.VisualStudio.TestTools.UnitTesting;
using Moq;
using OpenXMLSDK.Engine.Word;

namespace OpenXMLSDK.UnitTest.ReportEngine
{
    [TestClass]
    public class ReportEngineUnitTests
    {
        [TestMethod]
        public void Global_Generation()
        {
            ReportEngineTest.ReportEngine(string.Empty, string.Empty);
        }

        public void SaveDoc()
        {
            Assert.IsFalse(wdDoc != null);
        }
    }
}

