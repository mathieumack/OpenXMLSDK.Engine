using System;
using System.IO;
using System.Linq;
using System.Text;
using DO = DocumentFormat.OpenXml;
using DOP = DocumentFormat.OpenXml.Packaging;
using ReportEngine.Core.DataContext;
using ReportEngine.Core.Template;
using ReportEngine.Core.Template.Styles;

namespace OpenXMLSDK.Engine.Word.ReportEngine.Renders
{
    /// <summary>
    /// Table of contents extensions
    /// </summary>
    public static class TableOfContentsExtensions
    {
        /// <summary>
        /// Table of contents render
        /// </summary>
        /// <param name="tableOfContents"></param>
        /// <param name="documentPart"></param>
        /// <param name="context"></param>
        public static void Render(this TableOfContents tableOfContents, DOP.OpenXmlPart documentPart, ContextModel context)
        {
            AddToC(documentPart as DOP.MainDocumentPart, tableOfContents);
            AddToCStyles(documentPart as DOP.MainDocumentPart, tableOfContents, context);
        }

        /// <summary>
        /// Add the table of contents
        /// </summary>
        /// <param name="document"></param>
        /// <param name="tableOfContents"></param>
        public static void AddToC(DOP.MainDocumentPart documentPart, TableOfContents tableOfContents)
        {
            StringBuilder tocParameters = new StringBuilder();

            if (tableOfContents.StylesAndLevels.Any())
            {
                StringBuilder tocParametersFr = new StringBuilder();

                //English
                tocParameters.Append(@"TOC \h \z \t ");
                //French
                tocParametersFr.Append(@"TOC \h \z \t ");

                foreach (Tuple<string, string> styleAndLevel in tableOfContents.StylesAndLevels)
                {
                    //English
                    tocParameters.Append(string.Join(",", styleAndLevel.Item1, styleAndLevel.Item2));
                    tocParameters.Append(",");

                    //French
                    tocParametersFr.Append(string.Join(";", styleAndLevel.Item1, styleAndLevel.Item2));
                    tocParametersFr.Append(";");
                }

                tocParameters.Append(" ");
                tocParameters.Append(tocParametersFr);
            }
            else
                //default switches
                tocParameters.Append(@"TOC \o '1-3' \h \z \u");

            string xmlString =
            @"<w:sdt xmlns:w='http://schemas.openxmlformats.org/wordprocessingml/2006/main'>
              <w:sdtPr>
                <w:docPartObj>
                  <w:docPartGallery w:val='Table of Contents'/>
                  <w:docPartUnique/>
                </w:docPartObj>
              </w:sdtPr>
              <w:sdtEndPr>
                <w:rPr>
                 <w:rFonts w:asciiTheme='minorHAnsi' w:cstheme='minorBidi' w:eastAsiaTheme='minorHAnsi' w:hAnsiTheme='minorHAnsi'/>
                 <w:color w:val='auto'/>
                 <w:sz w:val='22'/>
                 <w:szCs w:val='22'/>
                 <w:lang w:eastAsia='en-US'/>
                </w:rPr>
              </w:sdtEndPr>
              <w:sdtContent>";

            if (!string.IsNullOrEmpty(tableOfContents.Title))
            {
                xmlString += @"
                <w:p>
                  <w:pPr>
                    <w:pStyle w:val='" + tableOfContents.TitleStyleId + @"'/>
                  </w:pPr>
                  <w:r>
                    <w:t>" + tableOfContents.Title + @"</w:t>
                  </w:r>
                </w:p>";
            }

            xmlString += @"
                <w:p>
                  <w:pPr>
                    <w:rPr>
                      <w:noProof/>
                    </w:rPr>
                  </w:pPr>
                  <w:pPr>
                    <w:tabs>
                        <w:tab w:val='right' w:leader='" + tableOfContents.LeaderCharValue.ToString() + @"'/>
                    </w:tabs>
                    <w:rPr>
                        <w:noProof/>
                    </w:rPr>
                  </w:pPr>
                  <w:r>
                    <w:fldChar w:fldCharType='begin' w:dirty='true'/>
                  </w:r>
                  <w:r>
                    <w:instrText xml:space='preserve'> " + tocParameters + @" </w:instrText>
                  </w:r>
                  <w:r>
                    <w:fldChar w:fldCharType='separate'/>
                  </w:r>
                </w:p>
                <w:p>
                  <w:r>
                    <w:rPr>
                      <w:b/>
                      <w:bCs/>
                      <w:noProof/>
                    </w:rPr>
                    <w:fldChar w:fldCharType='end'/>
                  </w:r>
                </w:p>
              </w:sdtContent>
            </w:sdt>";

            using (StreamWriter sw = new StreamWriter(new MemoryStream()))
            {
                sw.Write(xmlString);
                sw.Flush();
                sw.BaseStream.Seek(0, SeekOrigin.Begin);

                var re = DO.OpenXmlReader.Create(sw.BaseStream);

                re.Read();
                var oxe = re.LoadCurrentElement();
                documentPart.Document.Body.AppendChild(oxe);
                re.Close();
            }
        }

        /// <summary>
        /// Add styles for table of contents levels
        /// </summary>
        /// <param name="document"></param>
        /// <param name="tableOfContents"></param>
        /// <param name="context"></param>
        private static void AddToCStyles(DOP.MainDocumentPart document, TableOfContents tableOfContents, ContextModel context)
        {
            var stylesPart = document.StyleDefinitionsPart;
            if (tableOfContents.ToCStylesId.Any())
            {
                for (int i = 0; i < tableOfContents.ToCStylesId.Count; i++)
                {
                    var style = new Style()
                    {
                        StyleId = string.Concat("toc ", i + 1),
                        StyleBasedOn = tableOfContents.ToCStylesId[i],
                        CustomStyle = false
                    };
                    style.Render(stylesPart, context);
                }
            }
        }
    }
}
