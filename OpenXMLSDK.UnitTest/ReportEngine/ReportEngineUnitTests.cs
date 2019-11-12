using Microsoft.VisualStudio.TestTools.UnitTesting;
using OpenXMLSDK.Engine.Word;
using OpenXMLSDK.Engine.Word.Models;
using OpenXMLSDK.Engine.Word.Charts;
using OpenXMLSDK.Engine.Word.Paragraphs.Models;
using OpenXMLSDK.Engine.Word.Tables.Models;

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
            testSM.Val = (ShadingPatternValues)0;
            testSM.ThemeColor = (ThemeColorValues)0;
            testSM.ThemeFill = (ThemeColorValues)0;
            Assert.AreEqual(testSM.ThemeFillShade, testSM.ThemeFillTint);
            Assert.AreEqual(testSM.Color, testSM.Fill);
            Assert.AreEqual(testSM.ThemeShade, testSM.ThemeTint);
            Assert.AreEqual(testSM.ThemeColor, testSM.ThemeFill);
            Assert.AreEqual((ShadingPatternValues)0, testSM.Val);
        }

        [TestMethod]
        public void Test_Model_SpacingBetweenLinesModel()
        {
            SpacingBetweenLinesModel testSBLM = new SpacingBetweenLinesModel();
            testSBLM.After = "toto";
            testSBLM.Before = "toto";
            Assert.AreEqual(testSBLM.After, testSBLM.Before);
        }

        [TestMethod]
        public void Test_Charts()
        {
            ChartModelException testCharts = new ChartModelException("error");
            ChartModelException testCharts2 = new ChartModelException("message","error");
            ChartModelException testCharts3 = new ChartModelException("message", "error", null);
            Assert.AreEqual("error", testCharts.ErrorCode);
            Assert.IsNull(testCharts.InnerException);
        }

        [TestMethod]
        public void Test_Paragraphs_BM()
        {
            ParagraphBorderModel testPBM = new ParagraphBorderModel();
            Assert.AreEqual("FFFFFF", testPBM.Color);
            Assert.IsNull(testPBM.Size);
            Assert.AreEqual(BorderValues.None, testPBM.BorderValue);
        }

        [TestMethod]
        public void Test_Paragraphs_BMs()
        {
            ParagraphBordersModel testPBMs = new ParagraphBordersModel();
            testPBMs.BottomBorder = new ParagraphBorderModel();
            testPBMs.TopBorder = new ParagraphBorderModel();
            Assert.AreEqual(testPBMs.TopBorder.Color, testPBMs.BottomBorder.Color);
            Assert.AreEqual(testPBMs.TopBorder.Size, testPBMs.BottomBorder.Size);
            Assert.AreEqual(testPBMs.TopBorder.BorderValue, testPBMs.BottomBorder.BorderValue);

            testPBMs.LeftBorder = new ParagraphBorderModel();
            testPBMs.RightBorder = new ParagraphBorderModel();
            Assert.AreEqual(testPBMs.LeftBorder.Color, testPBMs.RightBorder.Color);
            Assert.AreEqual(testPBMs.LeftBorder.Size, testPBMs.RightBorder.Size);
            Assert.AreEqual(testPBMs.LeftBorder.BorderValue, testPBMs.RightBorder.BorderValue);
        }

        [TestMethod]
        public void Test_Paragraphs_PM()
        {
            ParagraphPropertiesModel testPM = new ParagraphPropertiesModel();
            testPM.NumberingProperties = new NumberingPropertiesModel();
            testPM.NumberingProperties.NumberingLevelReference = 1;
            testPM.NumberingProperties.NumberingId = 1;
            Assert.AreEqual(testPM.NumberingProperties.NumberingId, testPM.NumberingProperties.NumberingLevelReference);

            testPM.ParagraphStyleId = new ParagraphStyleIdModel();
            testPM.ParagraphStyleId.Val = "toto";
            Assert.AreEqual("toto", testPM.ParagraphStyleId.Val);

            testPM.SpacingBetweenLines = new SpacingBetweenLinesModel();
            testPM.SpacingBetweenLines.After = "toto";
            testPM.SpacingBetweenLines.Before = "toto";
            Assert.AreEqual(testPM.SpacingBetweenLines.After, testPM.SpacingBetweenLines.Before);

            testPM.ParagraphBorders = new ParagraphBordersModel();
            //Assert à mutualiser
        }

        [TestMethod]
        public void Test_Table()
        {
            TableBorderModel testTBM = new TableBorderModel();
            Assert.AreEqual(BorderValues.Single, testTBM.BorderValue);
            Assert.AreEqual(1, testTBM.Size);
            Assert.AreEqual("000000", testTBM.Color);
        }
    }
}

