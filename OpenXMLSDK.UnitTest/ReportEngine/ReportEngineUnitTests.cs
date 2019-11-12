using Microsoft.VisualStudio.TestTools.UnitTesting;
using OpenXMLSDK.Engine;
using OpenXMLSDK.Engine.ReportEngine.DataContext;
using Moq;


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

        [TestMethod]
        public void TestBaseModel()
        {
            StringModel aled = new StringModel();
            aled.Value = "aled";
            Assert.IsTrue(aled.Value == "aled");
        }
    }
}

