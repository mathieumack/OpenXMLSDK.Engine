using DocumentFormat.OpenXml.Packaging;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OpenXMLSDK.Engine.Word;
using System.Buffers.Text;
using System;
using System.IO;
using DocumentFormat.OpenXml.Validation;

namespace OpenXMLSDK.UnitTest.Managers.Word
{
    [TestClass]
    public class WordManagerNewTests
    {
        /// <summary>
        /// Test for method New
        /// </summary>
        [TestMethod]
        public void New()
        {
            // Arrange
            var wordManager = new WordManager();
            // Act
            var result = wordManager.New();

            // Assert
            Assert.AreEqual(true, result);

            // Save base64 file locally to test it:
            var path = Guid.NewGuid().ToString() + ".docx";

            wordManager.SaveDoc();
            var documentStream = wordManager.GetMemoryStream();
            var base64 = Convert.ToBase64String(documentStream.ToArray());
            File.WriteAllBytes(path, Convert.FromBase64String(base64));
            using (var wordDoc = WordprocessingDocument.Open(path, false))
            {
                var validator = new OpenXmlValidator();
                var errors = validator.Validate(wordDoc);
            }
        }
    }
}
