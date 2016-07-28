using System;
using System.IO;
using MvvX.Open_XML_SDK.Word;

namespace MvvX.Open_XML_SDK.TestConsole
{
    class Program
    {
        static void Main(string[] args)
        {
            var resourceName = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "App_Data", "Global.docx");

            using (var wordManager = new WordManager())
            {
                // TODO for debug : use your test file :
                wordManager.OpenDocFromTemplate(resourceName, Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "App_Data", "Results", "FinalDoc_Test_OrientationParagraph-" + DateTime.Now.ToFileTime() + ".docx"), true);

                wordManager.SetTextOnBookmark("Insert_Documents", "Hi !");

                wordManager.SaveDoc();
                wordManager.CloseDoc();
            }
        }
    }
}
