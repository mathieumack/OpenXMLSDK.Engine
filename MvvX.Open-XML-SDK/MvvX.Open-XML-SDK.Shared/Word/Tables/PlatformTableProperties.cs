using System;
using DocumentFormat.OpenXml.Wordprocessing;
using MvvX.Open_XML_SDK.Core.Word.Tables;
using DocumentFormat.OpenXml;

namespace MvvX.Open_XML_SDK.Shared.Word.Tables
{
    public class PlatformTableProperties : PlatformOpenXmlElement, ITableProperties
    {
        private readonly TableProperties tableProperties;

        public PlatformTableProperties(TableProperties tableProperties)
            : base(tableProperties)
        {
            this.tableProperties = tableProperties;
        }

        #region Interface

        private ITableBorders tableborder;
        public ITableBorders TableBorders
        {
            get
            {
                if (tableborder == null)
                    tableborder = PlatformTableBorders.New(tableProperties.TableBorders);

                return tableborder;
            }
        }

        public Core.Word.Tables.TableRowAlignmentValues? TableJustification
        {
            get
            {
                if (tableProperties.TableJustification == null || !tableProperties.TableJustification.Val.HasValue)
                    return null;
                else
                    return (Core.Word.Tables.TableRowAlignmentValues)(int)tableProperties.TableJustification.Val.Value;
            }
            set
            {
                if (value == null)
                    tableProperties.TableJustification = null;
                else
                {
                    if (tableProperties.TableJustification == null)
                        tableProperties.TableJustification = new TableJustification();
                    tableProperties.TableJustification.Val = (DocumentFormat.OpenXml.Wordprocessing.TableRowAlignmentValues)(int)value;
                }
            }
        }

        #endregion

        #region Static helpers methods

        public static PlatformTableProperties New()
        {
            return new PlatformTableProperties(new TableProperties());
        }

        #endregion
    }
}
