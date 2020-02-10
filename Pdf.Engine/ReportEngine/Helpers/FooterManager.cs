//using System;
//using RCT = ReportEngine.Core.Template;
//using RCD = ReportEngine.Core.DataContext;
//using iTextSharp.text.pdf;
//using iTextSharp.text;

//namespace Pdf.Engine.ReportEngine.Helpers
//{
//    internal class FooterManager
//    {
//        private PdfTemplate template;
//        private RCT.Footer footerModel;


//        public FooterManager(RCT.Footer footerModel)
//        {
//            this.footerModel = footerModel;
//        }

//        /// <summary>
//        /// Déclanche le pré-rendu de l'élément lors de l'ouverture du document
//        /// </summary>
//        /// <param name="Element"></param>
//        public void OnOpenDocumentRender(PdfWriter writer, Document document)
//        {
//            if (footerModel != null)
//                template = writer.DirectContent.CreateTemplate(500, footerModel.FontSize);
//        }

//        /// <summary>
//        /// Déclanche le rendu de l'élément lors de la fin d'une page
//        /// Les noeuds enfants du noeud footer ne sont pas gérés
//        /// </summary>
//        /// <param name="Element"></param>
//        public void OnEndPageRender(PdfWriter writer, Document document, EngineContext ctx, RCD.ContextModel Params)
//        {
//            if (footerModel != null && footerModel.PositionX.HasValue && footerModel.PositionY.HasValue)
//            {
//                if ((writer.PageNumber == 1 && !footerModel.ShowOnFirstPage))
//                    return;

//                BaseFont font = footerModel.GetBaseFont();

//                PdfContentByte cb = writer.DirectContent;
//                cb.SaveState();

//                float textBase = document.Bottom;
//                cb.BeginText();
//                cb.SetFontAndSize(font, footerModel.FontSize);

//                cb.SetColorFill(FontHelper.ConverPdfColorToColor(footerModel.FontColor));

//                float adjust = font.GetWidthPoint(string.Format(footerModel.Label, writer.PageNumber), footerModel.FontSize);
//                cb.SetTextMatrix(Utilities.MillimetersToPoints(footerModel.PositionX.Value), Utilities.MillimetersToPoints(footerModel.PositionY.Value));
//                cb.ShowText(string.Format(footerModel.Label, writer.PageNumber));
//                cb.EndText();

//                cb.AddTemplate(template, Utilities.MillimetersToPoints(footerModel.PositionX.Value) + adjust, Utilities.MillimetersToPoints(footerModel.PositionY.Value));
//                cb.RestoreState();

//                if (!string.IsNullOrEmpty(footerModel.Text))
//                {
//                    Params.ReplaceItem(footerModel);
//                    var paramText = Params.ReplaceText(footerModel.Text);
//                    ColumnText ct = new ColumnText(cb);
//                    var text = new Phrase(new Chunk(paramText, FontFactory.GetFont(FontFactory.HELVETICA, FontEncodings.CP1252, footerModel.FontSize, Font.NORMAL,
//                                                    new Color(footerModel.TextFontColor.Red, footerModel.TextFontColor.Green, footerModel.TextFontColor.Blue))));
//                    ct.SetSimpleColumn(text, footerModel.TextPositionX, footerModel.TextPositionY, footerModel.TextWidth, footerModel.TextHeight, 8, Element.ALIGN_LEFT);
//                    ct.Go();

//                    cb.BeginText();
//                    cb.SetFontAndSize(font, footerModel.FontSize);

//                    cb.SetColorFill(FontHelper.ConverPdfColorToColor(footerModel.TextFontColor));

//                    cb.SetTextMatrix(Utilities.MillimetersToPoints(footerModel.TextPositionX), Utilities.MillimetersToPoints(footerModel.TextPositionY));
//                    cb.SetLineWidth(100);
//                    cb.ShowText(footerModel.Text);
//                    cb.EndText();
//                }

//                if (footerModel.ChildImage != null)
//                    footerModel.ChildImage.Render(writer, document, ctx, Params);

//                if (footerModel.ChildRectangle != null)
//                {
//                    ctx.Parents.Add(footerModel);
//                    footerModel.ChildRectangle.Render(writer, document, ctx, Params);
//                }

//            }
//        }

//        /// <summary>
//        /// Déclanche le rendu de l'élément lors de la fin d'une page
//        /// </summary>
//        /// <param name="Element"></param>
//        public void OnCloseDocumentRender(PdfWriter writer, Document document)
//        {
//            if (template != null)
//            {
//                BaseFont font = footerModel.GetBaseFont();
//                template.BeginText();
//                template.SetFontAndSize(font, footerModel.FontSize);
//                if (footerModel.ShowTotalPageNumer)
//                {
//                    int pageNumber = writer.PageNumber - 1;
//                    template.ShowText(Convert.ToString(pageNumber));
//                }
//                template.EndText();
//            }
//        }
//    }
//}
