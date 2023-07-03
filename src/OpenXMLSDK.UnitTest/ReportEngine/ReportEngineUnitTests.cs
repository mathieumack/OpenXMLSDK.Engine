using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Collections.Generic;
using System;
using System.IO;
using System.Globalization;
using OpenXMLSDK.Engine.Word.ReportEngine;
using OpenXMLSDK.Engine.Word;

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
        public void Global_Generation_Stream()
        {
            // Configuration of sample data :
            var rootFolder = Path.Combine(Environment.CurrentDirectory, "Resources/MdFiles");
            var templatePath = Path.Combine(Environment.CurrentDirectory, "Resources/Dotx/sample.dotx");
            var outputPath = Path.Combine(Environment.CurrentDirectory, "Resources/Dotx/sample.docx");

            // Generate Word document :
            using var templateStream = File.OpenRead(templatePath);
            var markdownContent = File.ReadAllText(Path.Combine(rootFolder, "0.md"));

            // Launch transformation :
            var reports = new List<Report>();

            var culture = new CultureInfo("en-US");
            using var templateDocument = File.OpenRead(templatePath);

            var output = Stream.Null;
            using (var word = new WordManager())
            {
                if (templateDocument != null && templateDocument != Stream.Null)
                    word.OpenDocFromTemplate(templateDocument);

                // Append documentation :
                word.AppendSubDocument(reports, true, culture);

                word.SaveDoc();

                output = word.GetMemoryStream();
            }
        }
    }
}