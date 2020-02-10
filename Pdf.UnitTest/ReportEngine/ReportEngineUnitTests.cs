using System.Diagnostics;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace PdfSharp.UnitTest.ReportEngine
{
    [TestClass]
    public class ReportEngineUnitTests
    {
        [TestMethod]
        public void Global_Generation()
        {
            ReportEngineTest.ReportEngine(string.Empty, string.Empty);
        }
    }
}

