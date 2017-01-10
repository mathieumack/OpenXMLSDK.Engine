using DocumentFormat.OpenXml.Wordprocessing;
using MvvX.Plugins.OpenXMLSDK.Word.Tables;

namespace MvvX.Plugins.OpenXMLSDK.Platform.Word.Tables
{
    public class PlatformTableLayout : PlatformOpenXmlElement, ITableLayout
    {
        private readonly TableLayout xmlElement;

        public PlatformTableLayout()
            : this(new TableLayout())
        {
        }

        public PlatformTableLayout(TableLayout tableLayout)
            : base(tableLayout)
        {
            this.xmlElement = tableLayout;
        }

        #region Interface :

        public OpenXMLSDK.Word.Tables.TableLayoutValues? Type
        {
            get
            {
                if (xmlElement.Type.HasValue)
                    return (OpenXMLSDK.Word.Tables.TableLayoutValues)(int)xmlElement.Type.Value;
                else
                    return null;
            }

            set
            {
                if (!value.HasValue)
                    xmlElement.Type = null;
                else
                    xmlElement.Type = (DocumentFormat.OpenXml.Wordprocessing.TableLayoutValues)(int)value.Value;
            }
        }

        #endregion

        #region Static helpers methods

        public static PlatformTableLayout New(TableProperties tableProperties)
        {
            var xmlElement = CheckDescendantsOrAppendNewOne<TableLayout>(tableProperties);
            return new PlatformTableLayout(xmlElement);
        }

        #endregion
    }
}

