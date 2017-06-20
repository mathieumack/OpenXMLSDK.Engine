using System.IO;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using MvvX.Plugins.OpenXMLSDK.Word.ReportEngine.Models;

namespace MvvX.Plugins.OpenXMLSDK.Platform.Word.ReportEngine.TOC
{
    /// <summary>
    /// Table of contents extensions
    /// </summary>
    public static class TableOfContentsExtensions
    {
        public static void Render(this TableOfContents tableOfContents, WordprocessingDocument wdDoc)
        {
            //XElement firstPara = wdDoc.MainDocumentPart.GetXDocument().Descendants(W.p).FirstOrDefault();

            AddToc(wdDoc, @"TOC TOC \h \z \t Red;1", null, null);
        }

        public static void AddToc(WordprocessingDocument wdDoc, string switches, string title, int? rightTabPos)
        {

            if (title == null)
                title = "Contents";
            if (rightTabPos == null)
                rightTabPos = 9350;

            // {0} tocTitle (default = "Contents")
            // {1} rightTabPosition (default = 9350)
            // {2} switches

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
              <w:sdtContent>
                <w:p>
                  <w:pPr>
                    <w:pStyle w:val='TOCHeading'/>
                  </w:pPr>
                  <w:r>
                    <w:t>{0}</w:t>
                  </w:r>
                </w:p>
                <w:p>
                  <w:pPr>
                    <w:pStyle w:val='TOC1'/>
                    <w:tabs>
                      <w:tab w:val='right' w:leader='dot' w:pos='{1}'/>
                    </w:tabs>
                    <w:rPr>
                      <w:noProof/>
                    </w:rPr>
                  </w:pPr>
                  <w:r>
                    <w:fldChar w:fldCharType='begin' w:dirty='true'/>
                  </w:r>
                  <w:r>
                    <w:instrText xml:space='preserve'> {2} </w:instrText>
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

            XElement sdt = XElement.Parse(string.Format(xmlString, title, rightTabPos, switches));

            using (StreamWriter sw = new StreamWriter(new MemoryStream()))
            {
                sw.Write(sdt.ToString());
                sw.Flush();
                sw.BaseStream.Seek(0, SeekOrigin.Begin);

                OpenXmlReader re = OpenXmlReader.Create(sw.BaseStream);

                re.Read();
                OpenXmlElement oxe = re.LoadCurrentElement();
                wdDoc.MainDocumentPart.Document.Body.AppendChild(oxe);
                re.Close();
            }
        }
    }
}
