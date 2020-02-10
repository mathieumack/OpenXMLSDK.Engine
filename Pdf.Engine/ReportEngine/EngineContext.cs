using System.Collections.Generic;
using it = iTextSharp.text;
using itp = iTextSharp.text.pdf;
using Pdf.Engine.ReportEngine.Helpers;
using ReportEngine.Core.Template;
using ReportEngine.Core.Template.Text;
using System;

namespace Pdf.Engine.ReportEngine
{
    internal class EngineContext
    {
        // Contient la liste des objets IElementListener du document
        public IDictionary<BaseElement, it.IElementListener> IElementContainers { get; }

        public IDictionary<BaseElement, itp.PdfPCell> PCellContainers { get; }

        //public ChapterLevelCmpt ChapterLevels { get; set; }

        public List<BaseElement> Parents { get; set; }

        public EngineContext()
        {
            IElementContainers = new Dictionary<BaseElement, it.IElementListener>();
            PCellContainers = new Dictionary<BaseElement, itp.PdfPCell>();
            //ChapterLevels = new ChapterLevelCmpt();
            Parents = new List<BaseElement>();
        }
    }
}
