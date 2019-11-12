using Microsoft.VisualStudio.TestTools.UnitTesting;
using OpenXMLSDK.Engine.Word;
using OpenXMLSDK.Engine;
using Moq;
using System.IO;
using System;

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
        public void WordManagerOpenDocException()
        {
            WordManager manager = new WordManager();
            Exception expectedExcetpion = null;

            try
            {
                manager.OpenDoc("c", true);
            }
            catch (Exception ex)
            {
                expectedExcetpion = ex;
            }

            Assert.IsNotNull(expectedExcetpion);
        }

        [TestMethod]
        public void WordManagerOpenDocFromTemplateException()
        {
            WordManager manager = new WordManager();
            Exception expectedExcetpion = null;

            try
            {
                manager.OpenDocFromTemplate("c");
            }
            catch (Exception ex)
            {
                expectedExcetpion = ex;
            }

            Assert.IsNotNull(expectedExcetpion);
        }

        [TestMethod]
        public void WordManagerFindBookmarkException()
        {
            WordManager manager = new WordManager();
            Exception expectedExcetpion = null;

            try
            {
                manager.FindBookmark("c");
            }
            catch (Exception ex)
            {
                expectedExcetpion = ex;
            }

            Assert.IsNotNull(expectedExcetpion);
        }
    }
}

