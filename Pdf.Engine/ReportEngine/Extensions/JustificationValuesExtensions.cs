using System;
using System.Collections.Generic;
using System.Text;
using ReportEngine.Core.Template;

namespace Pdf.Engine.ReportEngine.Extensions
{
    public static class JustificationValuesExtensions
    {
        public static int ToPdfJustification(this JustificationValues justification)
        {
            switch(justification)
            {
                case JustificationValues.Left:
                    return 0;
                case JustificationValues.Start:
                    return 7;
                case JustificationValues.Center:
                    return 1;
                case JustificationValues.Right:
                    return 2;
                case JustificationValues.End:
                    return 0;
                case JustificationValues.Both:
                    return 3;
                default:
                    return 0;
            }
        }
    }
}