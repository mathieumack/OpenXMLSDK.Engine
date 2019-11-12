using Microsoft.VisualStudio.TestTools.UnitTesting;
using Moq;
using OpenXMLSDK.Engine.ReportEngine.DataContext;
using OpenXMLSDK.Engine.Word.ReportEngine.Models;
using System;
using System.Collections.Generic;
using System.Text;


namespace OpenXMLSDK.UnitTest {
    [TestClass]
    public class UnitTest2 {
        [TestMethod]
        public void MyUnitTest () {
            Mock<DateTimeModel> mockDatou = new Mock<DateTimeModel> ();

            DateTimeModel mockObject = mockDatou.Object;

            Assert.IsTrue (condition: mockObject.GetCurrentDate () == new DateTime (2019, 01, 01));
        }
    }
}