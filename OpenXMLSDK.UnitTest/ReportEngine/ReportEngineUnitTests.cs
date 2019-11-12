using Microsoft.VisualStudio.TestTools.UnitTesting;
using OpenXMLSDK;
using OpenXMLSDK.Engine.ReportEngine.DataContext;
using OpenXMLSDK.Engine.Word;
using OpenXMLSDK.Engine.Word.Models;
using System;

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
        public void TestUnit1()
        {
            string string1 = "Hey";
            ReportEngineTest.ReportEngine(string.Empty, string1);
        }
        /*
        [TestMethod]
        public void Global_Generation_with_documentName()
        {
            string string1 = "Hey";
            string string2 = "YES";
            ReportEngineTest.ReportEngine(string2, string1);
        }
        
        [TestMethod]
        public void closeDoc()
        {
            using (var word = new WordManager())
            {
                
                word.CloseDoc();
            }
        }*/
        [TestMethod]
        public void DateTimeModel_ValueDefault()
        {
            DateTimeModel DateTimeUT = new DateTimeModel();

            Assert.IsTrue(DateTimeUT.Value == DateTime.MinValue);
        }
        [TestMethod]
        public void DateTimeModel_Value()
        {
            DateTimeModel DateTimeUT = new DateTimeModel();
            DateTime date = new DateTime(2008, 5, 1, 8, 30, 52);
            DateTimeUT.Value = date;
            Assert.IsTrue(DateTimeUT.Value == date);
        }
        [TestMethod]
        public void DateTimeModel_RanderPattern()
        {
            DateTimeModel DateTimeUT = new DateTimeModel();
            DateTimeUT.RenderPattern = "TEST";
            Assert.IsTrue(DateTimeUT.RenderPattern == "TEST");

        }
        /*
        [TestMethod]
        public void DateTimeModel_Render(IFormatProvider test)
        {
            IFormatProvider formatProvider = test;
            DateTimeModel DateTimeUT = new DateTimeModel();
            DateTimeUT.RenderPattern = null;
            string haHa = DateTimeUT.Render(formatProvider);
            Assert.IsTrue(haHa == formatProvider.ToString());
        }
        */
        [TestMethod]
        public void ShadingModel_Color()
        {
            ShadingModel shadingUT = new ShadingModel();
            shadingUT.Color = "test";
            Assert.IsTrue(shadingUT.Color == "test");
        }
        [TestMethod]
        public void ShadingModel_Fill()
        {
            ShadingModel shadingUT = new ShadingModel();
            shadingUT.Fill = "test";
            Assert.IsTrue(shadingUT.Fill == "test");
        }
        [TestMethod]
        public void ShadingModel_ThemeColor()
        {
            ShadingModel shadingUT = new ShadingModel();
            ThemeColorValues value = new ThemeColorValues();
            if (shadingUT.ThemeColor.HasValue)
            {
                shadingUT.ThemeColor = value;
                Assert.IsTrue(shadingUT.ThemeColor == value);
            }
        }
        [TestMethod]
        public void ShadingModel_ThemeFill()
        {
            ShadingModel shadingUT = new ShadingModel();
            ThemeColorValues value = new ThemeColorValues();
            if (shadingUT.ThemeFill.HasValue)
            {
                shadingUT.ThemeFill = value;
                Assert.IsTrue(shadingUT.ThemeFill == value);
            }
        }
        [TestMethod]
        public void ShadingModel_ThemeFillShade()
        {
            ShadingModel shadingUT = new ShadingModel();
            shadingUT.ThemeFillShade = "test";
            Assert.IsTrue(shadingUT.ThemeFillShade == "test");
        }
        [TestMethod]
        public void ShadingModel_ThemeFillTint()
        {
            ShadingModel shadingUT = new ShadingModel();
            shadingUT.ThemeFillTint = "test";
            Assert.IsTrue(shadingUT.ThemeFillTint == "test");
        }
        [TestMethod]
        public void ShadingModel_ThemeShade()
        {
            ShadingModel shadingUT = new ShadingModel();
            shadingUT.ThemeShade = "test";
            Assert.IsTrue(shadingUT.ThemeShade == "test");
        }
        [TestMethod]
        public void ShadingModel_ThemeTint()
        {
            ShadingModel shadingUT = new ShadingModel();
            shadingUT.ThemeTint = "test";
            Assert.IsTrue(shadingUT.ThemeTint == "test");
        }
    }
}

