using System;
using System.Collections.Generic;
using System.Text;
using ReportEngine.Core.Template.Tables;

namespace Pdf.Engine.ReportEngine.Extensions
{
    public static class TableVerticalAlignmentValuesExtensions
    {
        public static int ToPdfVerticalAlignmentValues(this TableVerticalAlignmentValues tableVerticalAlignmentValues)
        {
            switch (tableVerticalAlignmentValues)
            {
                case TableVerticalAlignmentValues.Bottom:
                    return 6;
                case TableVerticalAlignmentValues.Center:
                    return 5;
                default:
                    return 4; // Top by default
            }
        }
    }
}
