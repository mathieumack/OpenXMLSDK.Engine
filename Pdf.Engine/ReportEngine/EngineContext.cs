using System;
using System.Collections.Generic;
using System.Text;
using it = iTextSharp.text;
using itp = iTextSharp.text.pdf;
using Pdf.Engine.ReportEngine.Helpers;
using ReportEngine.Core.Template;
using ReportEngine.Core.Template.Text;

namespace Pdf.Engine.ReportEngine
{
    internal class EngineContext
    {
        // Contient la liste des objets IElementListener du document
        public IDictionary<BaseElement, it.IElementListener> IElementContainers { get; }

        public IDictionary<BaseElement, itp.PdfPCell> PCellContainers { get; }

        public Stack<InheritProperties> Inherits { get; }

        //public ChapterLevelCmpt ChapterLevels { get; set; }

        public List<BaseElement> Parents { get; set; }

        public EngineContext()
        {
            IElementContainers = new Dictionary<BaseElement, it.IElementListener>();
            PCellContainers = new Dictionary<BaseElement, itp.PdfPCell>();
            //ChapterLevels = new ChapterLevelCmpt();
            Parents = new List<BaseElement>();
            Inherits = new Stack<InheritProperties>();
        }

        /// <summary>
        /// Update inherits
        /// </summary>
        /// <param name="document"></param>
        public void InitInherits(Document document)
        {
            Inherits.Push(new InheritProperties()
            {
                FontSize = document.DefaultFontSize
            });
        }

        /// <summary>
        /// Update inherits
        /// </summary>
        /// <param name="paragraph"></param>
        public void UpdateInherits(Paragraph paragraph)
        {
            var lastConfiguration = Inherits.Peek();

            Inherits.Push(new InheritProperties()
            {
                Shading = string.IsNullOrWhiteSpace(paragraph.Shading) ? lastConfiguration.Shading : paragraph.Shading,
                FontColor = lastConfiguration.FontColor,
                FontEncoding = lastConfiguration.FontEncoding,
                FontName = lastConfiguration.FontName,
                FontSize = lastConfiguration.FontSize
            });
        }

        public void EndInherits()
        {
            Inherits.Pop();
        }
    }
}
