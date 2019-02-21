using System;
using System.IO;
using System.Runtime.InteropServices;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using MvvX.Plugins.OpenXMLSDK.Word.ReportEngine.BatchModels;
using MvvX.Plugins.OpenXMLSDK.Word.ReportEngine.Models;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;

namespace MvvX.Plugins.OpenXMLSDK.Platform.Word.ReportEngine
{
    /// <summary>
    /// Image extention
    /// </summary>
    public static class ImageExtensions
    {
        /// <summary>
        /// Horizontal screen zoom level in dpi
        /// </summary>
        static readonly int Xdpi;

        /// <summary>
        /// Vertical screen zoom level in dpi
        /// </summary>
        static readonly int Ydpi;

        /// <summary>
        /// At first call, set Xdpi and Ydpi based on the screen zoom level
        /// </summary>
        static ImageExtensions()
        {
            System.Drawing.Graphics g = System.Drawing.Graphics.FromHwnd(IntPtr.Zero);
            IntPtr desktop = g.GetHdc();
            Xdpi = GetDeviceCaps(desktop, 88); // 88 for X axis
            Ydpi = GetDeviceCaps(desktop, 90); // 90 for Y axis
        }

        /// <summary>
        /// Retrieve device-specific information for the specified device
        /// </summary>
        /// <param name="hDC">Device context handle</param>
        /// <param name="nIndex">Index of the item to be returned</param>
        /// <returns></returns>
        [DllImport("gdi32.dll", CharSet = CharSet.Auto, SetLastError = true, ExactSpelling = true)]
        static extern int GetDeviceCaps(IntPtr hDC, int nIndex);

        /// <summary>
        /// Create the image
        /// </summary>
        /// <param name="image">Image model</param>
        /// <param name="parent">Container</param>
        /// <param name="context">Context</param>
        /// <param name="documentPart">MainDocumentPart</param>
        /// <returns></returns>
        public static OpenXmlElement Render(this Image image, OpenXmlElement parent, ContextModel context, OpenXmlPart documentPart)
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
                parent.AppendChild(result);
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
        private static OpenXmlElement CreateImage(ImagePart imagePart, Image model, OpenXmlPart mainDocumentPart)
        {
            string relationshipId = mainDocumentPart.GetIdOfPart(imagePart);

            long imageWidth;
            long imageHeight;

#if __WPF__ || __IOS__
            using (var bm = new System.Drawing.Bitmap(imagePart.GetStream()))
#endif
#if __ANDROID__
            using (var bm = Android.Graphics.BitmapFactory.DecodeStream(imagePart.GetStream()))
#endif
            {
                long bmWidth = bm.Width;
                long bmHeight = bm.Height;

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

#if __WPF__ || __IOS__
                // In case the image dpi is changed based on the screen zoom level when creating the bitmap
                // Revert the bitmap dpi to the default one (96), which represents the 100% zoom level (no zoom) on Windows
                // As the dpi is not always 96 even if not altered by the zoom level, don't touch (=> if not equal to the current zoom level)
                var horizontalResolution = bm.HorizontalResolution;
                if (Xdpi != 96 && Xdpi == bm.HorizontalResolution)
                    horizontalResolution = 96;

                var verticalResolution = bm.VerticalResolution;
                if (Ydpi != 96 && Ydpi == bm.VerticalResolution)
                    verticalResolution = 96;

                imageWidth = bmWidth * (long)(914400 / horizontalResolution);
                imageHeight = bmHeight * (long)(914400 / verticalResolution);
#endif
#if __ANDROID__
                // TODO : Check this method
                imageWidth = bmWidth * (long)((float)914400 / (long)bm.Density);
                imageHeight = bmHeight * (long)((float)914400 / (long)bm.Density);
#endif
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
                                             Name = "New Bitmap Image.jpg"
                                         },
                                         new PIC.NonVisualPictureDrawingProperties()),
                                     new PIC.BlipFill(
                                         new A.Blip()
                                         {
                                             Embed = relationshipId,
                                             CompressionState =
                                             A.BlipCompressionValues.Print
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

            result.Append(new DocumentFormat.OpenXml.Wordprocessing.Drawing(
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
                             Name = "Picture 1"
                         },
                         new DW.NonVisualGraphicFrameDrawingProperties(graphicFrameLocks),
                         graphic
                     )
                     {
                         DistanceFromTop = 0U,
                         DistanceFromBottom = 0U,
                         DistanceFromLeft = 0U,
                         DistanceFromRight = 0U
                     }));
            return result;
        }
    }
}
