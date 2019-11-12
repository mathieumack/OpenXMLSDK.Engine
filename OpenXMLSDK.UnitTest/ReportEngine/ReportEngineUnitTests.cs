using Microsoft.VisualStudio.TestTools.UnitTesting;
using OpenXMLSDK.Engine.Word;
using OpenXMLSDK.Engine.Word.Models;

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
        public void Test_Model_GridSpanModel()
        {
            GridSpanModel testGDM = new GridSpanModel();
            testGDM.Val = 1;
            Assert.AreEqual(1, testGDM.Val);
        }

        [TestMethod]
        public void Test_Model_NumberingPropertiesModel()
        {
            NumberingPropertiesModel testNPM = new NumberingPropertiesModel();
            testNPM.NumberingLevelReference = 1;
            testNPM.NumberingId = 1;
            Assert.AreEqual(testNPM.NumberingId, testNPM.NumberingLevelReference);
        }

        [TestMethod]
        public void Test_Model_RunFontsModel()
        {
            RunFontsModel testRFM = new RunFontsModel();
            testRFM.Ascii = "toto";
            testRFM.ComplexScript = "toto";
            testRFM.EastAsia = "toto";
            testRFM.HighAnsi = "toto";
            Assert.AreEqual(testRFM.Ascii, testRFM.ComplexScript);
            Assert.AreEqual(testRFM.EastAsia, testRFM.HighAnsi);
        }

        [TestMethod]
        public void Test_Model_RunPropertiesModel()
        {
            RunPropertiesModel testRPM = new RunPropertiesModel();
            testRPM.Bold = true;
            testRPM.Italic = true;
            testRPM.Color = "color";
            testRPM.FontSize = "color";
            testRPM.RunFonts = new RunFontsModel();
            Assert.AreEqual(testRPM.Bold, testRPM.Italic);
            Assert.AreEqual(testRPM.Color, testRPM.FontSize);
            Assert.IsNotNull(testRPM.RunFonts);
        }

        [TestMethod]
        public void Test_Model_ShadingModel()
        {
            ShadingModel testSM = new ShadingModel();
            testSM.ThemeFillShade = "toto";
            testSM.ThemeFillTint = "toto";
            testSM.Color = "color";
            testSM.Fill = "color";
            testSM.ThemeShade = "toto";
            testSM.ThemeTint = "toto";
            //testSM.Val = 0;
            testSM.ThemeColor = (ThemeColorValues)0;
            testSM.ThemeFill = (ThemeColorValues)0;
            Assert.AreEqual(testSM.ThemeFillShade, testSM.ThemeFillTint);
            Assert.AreEqual(testSM.Color, testSM.Fill);
            Assert.AreEqual(testSM.ThemeShade, testSM.ThemeTint);
            Assert.AreEqual(testSM.ThemeColor, testSM.ThemeFill);
            //Assert.AreEqual(0, testSM.ThemeFill);
        }
    }
}

