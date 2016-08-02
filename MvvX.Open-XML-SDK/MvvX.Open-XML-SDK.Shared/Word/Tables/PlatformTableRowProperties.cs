using System;
using System.Collections.Generic;
using System.Text;
using DocumentFormat.OpenXml.Wordprocessing;
using MvvX.Open_XML_SDK.Core.Word.Bases;

namespace MvvX.Open_XML_SDK.Shared.Word.Tables
{
    public class PlatformTableRowProperties : PlatformOpenXmlElement
    {
        public TableRowProperties tableRowProperties { get; set; }

        public PlatformTableRowProperties(TableRowProperties tableRowProperties)
            : base(tableRowProperties)
        {
            this.tableRowProperties = tableRowProperties;
        }
    }
}
