using System;
using System.Globalization;
using System.IO;
using Newtonsoft.Json;
using OpenXMLSDK.Engine.Word;
using ReportEngine.Core.DataContext;
using ReportEngine.Core.Template;
using SampleTest.Content;

namespace OpenXMLSDK.UnitTest.ReportEngine
{
    public static class ReportEngineTest
    {
        public static void ReportEngine(string filePath, string documentName)
        {
            // Debut test report engine
            byte[] res = null;
            var word = new WordManager();
            {
                JsonContextConverter[] converters = { new JsonContextConverter() };

                if (string.IsNullOrWhiteSpace(filePath))
                {
                    if (string.IsNullOrWhiteSpace(documentName))
                        documentName = "ExampleDocument.docx";

                    var template = ReportEngineSample.GetTemplateDocument();
                    var templateJson = JsonConvert.SerializeObject(template);
                    var templateUnserialized = JsonConvert.DeserializeObject<Document>(templateJson, new JsonSerializerSettings() { Converters = converters });

                    var context = ReportEngineSample.GetContext();
                    var contextJson = JsonConvert.SerializeObject(context);
                    var contextUnserialized = JsonConvert.DeserializeObject<ContextModel>(contextJson, new JsonSerializerSettings() { Converters = converters });
                     
                    res = word.GenerateReport(templateUnserialized, contextUnserialized, new CultureInfo("en-US"));
                }
                else
                {
                    if (string.IsNullOrWhiteSpace(documentName))
                        documentName = "ExampleDocument.docx";
                    if (!documentName.EndsWith(".docx"))
                        documentName = string.Concat(documentName, ".docx");

                    var stream = File.ReadAllText(filePath);
                    var report = JsonConvert.DeserializeObject<Report>(stream, new JsonSerializerSettings() { Converters = converters });

                    res = word.GenerateReport(report.Document, report.ContextModel, new CultureInfo("en-US"));
                }
            }

            // test ecriture fichier
            File.WriteAllBytes(documentName, res);
        }

        public static void Test()
        {
            Console.WriteLine("Enter the path of your Json file, press enter for an example");
            var filePath = Console.ReadLine();
            var documentName = string.Empty;
            if (!string.IsNullOrWhiteSpace(filePath))
            {
                Console.WriteLine("Enter document name");
                documentName = Console.ReadLine();
            }

            Console.WriteLine("Generation in progress");
            ReportEngine(filePath, documentName);
        }
    }
}