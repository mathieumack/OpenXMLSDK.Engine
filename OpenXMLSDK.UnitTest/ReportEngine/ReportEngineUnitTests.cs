using Microsoft.VisualStudio.TestTools.UnitTesting;
using Moq;
using OpenXMLSDK.Engine.Word;
using OpenXMLSDK.Engine.ReportEngine.DataContext;

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
        public void WordManager_CloseDoc()
        {
            SaveDoc
        }

        [TestMethod]
        public void DateTime()
        {
            DateTimeModel DateTime = new DateTimeModel();
        }

        [TestMethod]
        public void RenderPattern()
        {
            DateTimeModel RenderPattern = new DateTimeModel();
        }

        


    }
}

