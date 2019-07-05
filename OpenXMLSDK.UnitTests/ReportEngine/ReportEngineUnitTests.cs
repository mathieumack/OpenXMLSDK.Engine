using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace OpenXMLSDK.UnitTests.ReportEngine
{
    [TestClass]
    public class ReportEngineUnitTests
    {
        [TestMethod]
        public void Global_Generation()
        {
            ReportEngineTest.ReportEngine(string.Empty, string.Empty);
        }
    }
}
