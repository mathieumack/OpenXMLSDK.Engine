using Microsoft.VisualStudio.TestTools.UnitTesting;
using Moq;
using OpenXMLSDK.Engine.Word.ReportEngine.Models;
using System;
using System.Collections.Generic;
using System.Text;


[TestClass]
public class UnitTest1 {
    [TestMethod]
    public void MyUnitTest () {
        Mock<AlternateRowCellConfiguration> mockAlternate = new Mock<AlternateRowCellConfiguration> ();

        AlternateRowCellConfiguration mockObject = mockAlternate.Object;

        Assert.Fail ();
    }


}



/*
namespace OpenXMLSDK.UnitTest {
    class TestClass {

    }
} */
