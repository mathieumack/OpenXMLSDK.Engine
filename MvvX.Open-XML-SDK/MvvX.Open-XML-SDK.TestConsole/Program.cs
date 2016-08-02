using MvvX.Open_XML_SDK.PCL;
using MvvX.Open_XML_SDK.Word;

namespace MvvX.Open_XML_SDK.TestConsole
{
    class Program
    {
        static void Main(string[] args)
        {
            using (var wordManager = new WordManager())
            {
                WordReport wordReport = new WordReport(wordManager);
                wordReport.GenerateReport();
            }

        }
    }
}
