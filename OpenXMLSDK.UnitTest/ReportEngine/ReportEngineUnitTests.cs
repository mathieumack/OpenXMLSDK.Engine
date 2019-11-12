using Microsoft.VisualStudio.TestTools.UnitTesting;
using OpenXMLSDK.Engine;
using Moq;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OpenXMLSDK.Engine.ReportEngine.DataContext;
using OpenXMLSDK.Engine.Word.ReportEngine;
using OpenXMLSDK.Engine.Word.ReportEngine.Renders;
using OpenXMLSDK.Engine.Word.Models;
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
		public void GridSpanModel()
		{
			var Model = new GridSpanModel();
			Assert.IsTrue(Model != null);
		}

		[TestMethod]
		public void NumberingProperties()
		{
			var properties = new NumberingPropertiesModel();
			properties.NumberingLevelReference = 2;
			Assert.IsTrue(properties.NumberingLevelReference == 2);
		}

		[TestMethod]
		public void NumberingProperties2()
		{
			var properties = new NumberingPropertiesModel();
			properties.NumberingId = 8;
			Assert.IsTrue(properties.NumberingId == 8);
		}

		[TestMethod]
		public void RunFontsModel()
		{
			var models = new RunFontsModel();
			models.Ascii = "oui, c'est la vie!";
			Assert.IsTrue(models.Ascii == "oui, c'est la vie!");
		}

		[TestMethod]
		public void RunFontsModel2()
		{
			var models = new RunFontsModel();
			models.ComplexScript = "cool";
			Assert.IsTrue(models.ComplexScript == "cool");
		}

		[TestMethod]
		public void RunFontsModel3()
		{
			var models = new RunFontsModel();
			models.EastAsia = "c'est ça!";
			Assert.IsTrue(models.EastAsia == "c'est ça!");
		}

		[TestMethod]
		public void RunFontsModel4()
		{
			var models = new RunFontsModel();
			models.HighAnsi = "voyez pltuôt!";
			Assert.IsTrue(models.HighAnsi == "voyez pltuôt!");
		}

		[TestMethod]
		public void RunPropertiesModel()
		{
			var models = new RunPropertiesModel();
			models.Bold = true;
			Assert.IsTrue(models.Bold.Value);
			models.Bold = false;
			Assert.IsFalse(models.Bold.Value);
		}

		[TestMethod]
		public void RunPropertiesMode2()
		{
			var models = new RunPropertiesModel();
			models.Italic = true;
			Assert.IsTrue(models.Italic.Value);
			models.Italic = false;
			Assert.IsFalse(models.Italic.Value);
		}

		[TestMethod]
		public void RunPropertiesModel3()
		{
			var models = new RunPropertiesModel();
			models.Color = "red";
			Assert.IsTrue(models.Color=="red");
		}

		[TestMethod]
		public void RunPropertiesModel4()
		{
			var models = new RunPropertiesModel();
			models.FontSize = "15px";
			Assert.IsTrue(models.FontSize == "15px");
		}

		[TestMethod]
		public void DateTimeModel()
		{
			var dateTime = new DateTimeModel();
			Assert.IsTrue(dateTime != null);
			Assert.IsTrue(dateTime.Value == DateTime.MinValue);
		}

		[TestMethod]
		public void TableCellPropertiesModel()
		{
			var tableProperties = new TableCellPropertiesModel();
			tableProperties.Fusion = true;
			Assert.IsTrue(tableProperties.Fusion == true);
		}

		[TestMethod]
		public void TableCellPropertiesModel2()
		{
			var tableProperties = new TableCellPropertiesModel();
			tableProperties.FusionChild = true;
			Assert.IsTrue(tableProperties.FusionChild == true);
		}

		[TestMethod]
		public void TableCellPropertiesModel3()
		{
			var tableProperties = new TableCellPropertiesModel();
			tableProperties.Height = "12px";
			Assert.IsTrue(tableProperties.Height == "12px");
		}

		[TestMethod]
		public void TableCellPropertiesModel4()
		{
			var tableProperties = new TableCellPropertiesModel();
			tableProperties.ParagraphSolidarity = true;
			Assert.IsTrue(tableProperties.ParagraphSolidarity == true);
		}
	}
}

