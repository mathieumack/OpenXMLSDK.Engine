using System;

namespace ReportEngine.Core.Template.ExtendedModels
{
    /// <summary>
    /// Borders position
    /// </summary>
    [Flags]
    public enum BorderPositions
    {
        /// <summary>
        /// No border
        /// </summary>
        NONE = 0,
        /// <summary>
        /// border on top
        /// </summary>
        TOP = 1,
        /// <summary>
        /// border on bottom
        /// </summary>
        BOTTOM = 2,
        /// <summary>
        /// border on the left
        /// </summary>
        LEFT = 4,
        /// <summary>
        ///  border on the right
        /// </summary>
        RIGHT = 8,
        /// <summary>
        /// For table border only : horizontal border inside table (between rows)
        /// </summary>
        INSIDEHORIZONTAL = 16,
        /// <summary>
        /// For table border only : vertical border inside table (between columns)
        /// </summary>
        INSIDEVERTICAL = 32
    }
}
