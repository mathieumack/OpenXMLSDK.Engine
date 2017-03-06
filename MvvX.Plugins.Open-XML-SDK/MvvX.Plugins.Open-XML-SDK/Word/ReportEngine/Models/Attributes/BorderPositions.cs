using System;

namespace MvvX.Plugins.OpenXMLSDK.Word.ReportEngine.Models.Attributes
{
    [Flags]
    public enum BorderPositions
    {
        NONE = 0,
        TOP = 1,
        BOTTOM = 2,
        LEFT = 4,
        RIGHT = 8,
        INSIDEHORIZONTAL = 16,
        INSIDEVERTICAL = 32
    }
}
