﻿using System;
using System.IO;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OpenXMLSDK.Engine.ReportEngine.DataContext;
using SixLabors.ImageSharp.Metadata;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;

namespace OpenXMLSDK.Engine.Word.ReportEngine.Renders
{
    /// <summary>
    /// Image extention
    /// </summary>
    public static class ImageExtensions
    {
        /// <summary>
        /// Create the image
        /// </summary>
        /// <param name="image">Image model</param>
        /// <param name="parent">Container</param>
        /// <param name="context">Context</param>
        /// <param name="documentPart">MainDocumentPart</param>
        /// <returns></returns>
        public static OpenXmlElement Render(this Models.Image image, OpenXmlElement parent, ContextModel context, OpenXmlPart documentPart)
        {
            context.ReplaceItem(image);
            ImagePart imagePart;
            if (documentPart is MainDocumentPart)
                imagePart = (documentPart as MainDocumentPart).AddImagePart((ImagePartType)(int)image.ImagePartType);
            else if (documentPart is HeaderPart)
                imagePart = (documentPart as HeaderPart).AddImagePart((ImagePartType)(int)image.ImagePartType);
            else if (documentPart is FooterPart)
                imagePart = (documentPart as FooterPart).AddImagePart((ImagePartType)(int)image.ImagePartType);
            else
                return null;

            bool isNotEmpty = false;
            if (image.Content != null && image.Content.Length > 0)
            {
                using (MemoryStream stream = new MemoryStream(image.Content))
                {
                    imagePart.FeedData(stream);
                }
                isNotEmpty = true;
            }
            else if (!string.IsNullOrWhiteSpace(image.Path))
            {
                using (FileStream stream = new FileStream(image.Path, FileMode.Open))
                {
                    imagePart.FeedData(stream);
                }
                isNotEmpty = true;
            }
            if (isNotEmpty)
            {
                OpenXmlElement result = CreateImage(imagePart, image, documentPart);
                parent.Append(result);

                return result;
            }
            else
            {
                return null;
            }
        }

        /// <summary>
        /// Create the image to integrate
        /// </summary>
        /// <param name="imagePart"></param>
        /// <param name="model"></param>
        /// <param name="mainDocumentPart"></param>
        /// <returns></returns>
        private static OpenXmlElement CreateImage(ImagePart imagePart, Models.Image model, OpenXmlPart mainDocumentPart)
        {
            string relationshipId = mainDocumentPart.GetIdOfPart(imagePart);

            long imageWidth;
            long imageHeight;

            var hyperlink = GetHyperlinkRelationShip(mainDocumentPart, model);

            using (var image = SixLabors.ImageSharp.Image.Load(imagePart.GetStream()))
            {
                long bmWidth = image.Width;
                long bmHeight = image.Height;

                // Resize width
                if (model.Width.HasValue)
                {
                    long ratio = model.Width.Value * 100L / bmWidth;

                    bmWidth = (long)(bmWidth * (ratio / 100D));
                    bmHeight = (long)(bmHeight * (ratio / 100D));
                }

                // Resize width if too big
                if (model.MaxWidth.HasValue && model.MaxWidth.Value < bmWidth)
                {
                    long ratio = model.MaxWidth.Value * 100L / bmWidth;

                    bmWidth = (long)(bmWidth * (ratio / 100D));
                    bmHeight = (long)(bmHeight * (ratio / 100D));
                }

                // Resize height
                if (model.Height.HasValue)
                {
                    long ratio = model.Height.Value * 100L / bmHeight;

                    bmWidth = (long)(bmWidth * (ratio / 100D));
                    bmHeight = (long)(bmHeight * (ratio / 100D));
                }

                // Resize height if too big
                if (model.MaxHeight.HasValue && model.MaxHeight.Value < bmHeight)
                {
                    long ratio = model.MaxHeight.Value * 100L / bmHeight;

                    bmWidth = (long)(bmWidth * (ratio / 100D));
                    bmHeight = (long)(bmHeight * (ratio / 100D));
                }

                var xResolution = image.Metadata.HorizontalResolution;
                var yResolution = image.Metadata.VerticalResolution;

                // The resolution may come in differents units, convert it to pixels per inch
                if (image.Metadata.ResolutionUnits == PixelResolutionUnit.PixelsPerMeter)
                {
                    xResolution *= 0.0254;
                    yResolution *= 0.0254;
                }
                else if (image.Metadata.ResolutionUnits == PixelResolutionUnit.PixelsPerCentimeter)
                {
                    xResolution *= 2.54;
                    yResolution *= 2.54;
                }

                imageWidth = bmWidth * (long)(914400 / xResolution);
                imageHeight = bmHeight * (long)(914400 / yResolution);
            }

            var result = new Run();

            var runProperties = new RunProperties();
            runProperties.AppendChild(new NoProof());
            result.AppendChild(runProperties);

            var graphicFrameLocks = new A.GraphicFrameLocks() { NoChangeAspect = true };
            graphicFrameLocks.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");
            var picture = new PIC.Picture(
                                     new PIC.NonVisualPictureProperties(
                                         new PIC.NonVisualDrawingProperties()
                                         {
                                             Id = 0U,
                                             Name = "New Bitmap Image.jpg",
                                             HyperlinkOnClick = GetHyperlinkOnClick(hyperlink),
                                         },
                                         new PIC.NonVisualPictureDrawingProperties()),
                                     new PIC.BlipFill(
                                         new A.Blip()
                                         {
                                             Embed = relationshipId,
                                             CompressionState = A.BlipCompressionValues.Print,
                                         },
                                         new A.Stretch(new A.FillRectangle())),
                                     new PIC.ShapeProperties(
                                         new A.Transform2D(
                                             new A.Offset() { X = 0L, Y = 0L },
                                             new A.Extents() { Cx = imageWidth, Cy = imageHeight }),
                                         new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle }));
            picture.AddNamespaceDeclaration("pic", "http://schemas.openxmlformats.org/drawingml/2006/picture");
            var graphic = new A.Graphic(
                             new A.GraphicData(
                                 picture
                             )
                             { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" });
            graphic.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");

            var drawing = new Drawing(
                     new DW.Inline(
                         new DW.Extent() { Cx = imageWidth, Cy = imageHeight },
                         new DW.EffectExtent()
                         {
                             LeftEdge = 0L,
                             TopEdge = 0L,
                             RightEdge = 0L,
                             BottomEdge = 0L
                         },
                         new DW.DocProperties()
                         {
                             Id = 1U,
                             Name = "Picture 1",
                             HyperlinkOnClick = GetHyperlinkOnClick(hyperlink),
                         },
                         new DW.NonVisualGraphicFrameDrawingProperties(graphicFrameLocks),
                         graphic
                     )
                     {
                         DistanceFromTop = 0U,
                         DistanceFromBottom = 0U,
                         DistanceFromLeft = 0U,
                         DistanceFromRight = 0U
                     });
            result.Append(drawing);

            return result;
        }

        private static HyperlinkRelationship GetHyperlinkRelationShip(OpenXmlPart part, Models.Image model)
        {
            if (model.Hyperlink != null)
                return part.AddHyperlinkRelationship(new Uri(model.Hyperlink.HyperlinkUrl), model.Hyperlink.IsExternalUrl);

            return null;
        }

        private static A.HyperlinkOnClick GetHyperlinkOnClick(HyperlinkRelationship hyperlink)
        {
            if (hyperlink != null && !string.IsNullOrEmpty(hyperlink.Id))
                return new A.HyperlinkOnClick { Id = hyperlink.Id };

            return null;
        }
    }
}
