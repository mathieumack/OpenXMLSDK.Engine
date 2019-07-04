using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace MvvX.Plugins.Open_XML_SDK.UnitTests
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
