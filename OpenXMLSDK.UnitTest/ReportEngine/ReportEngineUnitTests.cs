using Microsoft.VisualStudio.TestTools.UnitTesting;
using OpenXMLSDK.Engine.Word;
using System;
using System.IO;

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

        [TestMethod]
        public void WordManagerGetBookmarkException()
        {
            WordManager manager = new WordManager();
            Exception expectedExcetpion = null;

            try
            {
                manager.GetBookmarks();
            }
            catch (Exception ex)
            {
                expectedExcetpion = ex;
            }

            Assert.IsNotNull(expectedExcetpion);
        }

        [TestMethod]
        public void WordManagerOpenDocOpenningPdf()
        {
            WordManager manager = new WordManager();
            string file_path = @"c:\temp\test_unit.pdf";
            File.Create(file_path);
            Assert.IsFalse(manager.OpenDoc(file_path, true));
        }

        [TestMethod]
        public void WordManagerOpenDocOpenningTxt()
        {
            WordManager manager = new WordManager();
            string file_path = @"c:\temp\test_unit.txt";
            File.Create(file_path);
            Assert.IsFalse(manager.OpenDoc(file_path, true));
        }

        [TestMethod]
        public void WordManagerOpenDocOpenningZip()
        {
            WordManager manager = new WordManager();
            string file_path = @"c:\temp\test_unit.zip";
            File.Create(file_path);
            Assert.IsFalse(manager.OpenDoc(file_path, true));
        }
    }
}

