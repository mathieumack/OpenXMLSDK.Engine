using System;
using System.IO;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newtonsoft.Json;
using OpenXMLSDK.Engine.Validation;

namespace OpenXMLSDK.UnitTest.ReportEngine
{
    [TestClass]
    public class ReportEngineUnitTests
    {
        [TestMethod]
        public void Global_Generation()
        {
            ReportEngineTest.ReportEngine(string.Empty, string.Empty, false);
        }

        [TestMethod]
        public void ValidationTool()
        {
            var filePath = Guid.NewGuid().ToString()  + ".docx"; // @"c:\temp\OnSite_TechRpt_WO-06844979_200604-112127.docx"

            ReportEngineTest.ReportEngine(string.Empty, filePath, false);

            var errors = OpenXMLValidator.ValidateWordDocument(filePath);
            var savedErrors = errors.Select(e => new
            {
                e.Id,
                ErrorType = e.ErrorType.ToString(),
                e.Description,
                e.Path
            }).ToList();

            File.WriteAllText(filePath + ".json", JsonConvert.SerializeObject(savedErrors));
            Assert.AreEqual(0, savedErrors.Count);
        }
    }
}