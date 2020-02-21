using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace OpenXMLSDK.UnitTest.ReportEngine
{
    [TestClass]
    public class ReportEngineUnitTests
    {
        [TestMethod]
        public void Global_Generation()
        {
            ReportEngineTest.ReportEngine(string.Empty, string.Empty);
            //ReportEngineTest.ReportEngine(@"C:\Users\leyvraz\Desktop\reportContext.json", "technical_lelvel_test", true);
        }
    }
}

