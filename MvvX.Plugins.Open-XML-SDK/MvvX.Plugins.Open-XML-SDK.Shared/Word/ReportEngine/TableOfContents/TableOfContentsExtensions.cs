using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using MvvX.Plugins.OpenXMLSDK.Word.ReportEngine.Models;

namespace MvvX.Plugins.OpenXMLSDK.Platform.Word.ReportEngine

{
    /// <summary>
    /// Table of contents extensions
    /// </summary>
    public static class TableOfContentsExtensions
    {
        public static void Render(this TableOfContents tableOfContents, WordprocessingDocument wdDoc)
        {
            //XElement firstPara = wdDoc.MainDocumentPart.GetXDocument().Descendants(W.p).FirstOrDefault();

            AddToc(wdDoc, tableOfContents);
        }

        public static void AddToc(WordprocessingDocument wdDoc, TableOfContents tableOfContents)
        {
            //default switches
            string switches = @"TOC \o '1-3' \h \z \u";
            if (tableOfContents.StylesAndLevels.Any())
            {
                switches = @"TOC \h \z \t ";
                foreach (Tuple<string, string> styleAndLevel in tableOfContents.StylesAndLevels)
                {
                    switches += styleAndLevel.Item1 + ";" + styleAndLevel.Item2 + ";";
                }
            }

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
                    <w:instrText xml:space='preserve'> " + switches + @" </w:instrText>
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

                OpenXmlReader re = OpenXmlReader.Create(sw.BaseStream);

                re.Read();
                OpenXmlElement oxe = re.LoadCurrentElement();
                wdDoc.MainDocumentPart.Document.Body.AppendChild(oxe);
                re.Close();
            }
        }
    }
}
