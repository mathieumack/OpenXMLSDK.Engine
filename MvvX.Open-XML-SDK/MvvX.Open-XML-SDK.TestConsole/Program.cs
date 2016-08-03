using System;
using System.IO;
using MvvX.Open_XML_SDK.Core.Word.Images;
using MvvX.Open_XML_SDK.Word;

namespace MvvX.Open_XML_SDK.TestConsole
{
    class Program
    {
        static void Main(string[] args)
        {
            var resourceName = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "App_Data", "Global.docx");
            var imagePath = @"C:\temp\circle.png";

            using (var wordManager = new WordManager())
            {
                // TODO for debug : use your test file :
                wordManager.OpenDocFromTemplate(resourceName, Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "App_Data", "Results", "FinalDoc_Test_OrientationParagraph-" + DateTime.Now.ToFileTime() + ".docx"), true);

                //wordManager.SetTextOnBookmark("Insert_Documents", "Hi !");
                wordManager.InsertPictureToBookmark("Insert_Documents", imagePath, ImageType.Png);
                wordManager.SaveDoc();
                wordManager.CloseDoc();
            }
        }
    }
}
