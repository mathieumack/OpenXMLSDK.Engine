namespace ReportEngine.Core.Template.ExtendedModels
{
    /// <summary>
    /// Model class for a border
    /// </summary>
    public class BorderModel
    {
        /// <summary>
        /// Define borders displayed : 0 = None, 1 = Top, 2 = Bottom, 4 = Left, 8 = Right
        /// (it's a flagged enum : 3 = Top and Bottom, 7 = Top and bottom and left, etc...)
        /// </summary>
        public BorderPositions BorderPositions { get; set; }

        /// <summary>
        /// border width size in eighths of a point 
        /// </summary>
        public uint BorderWidth { get; set; }

        /// <summary>
        /// If true : all borders use BorderWidth property
        /// If false : each border use it's own borderwidth property
        /// </summary>
        public bool UseVariableBorders { get; set; }

        /// <summary>
        /// Left border width (used only if UseVariableBorders = true)
        /// </summary>
        public uint BorderWidthLeft { get; set; }

        /// <summary>
        /// Right border width (used only if UseVariableBorders = true)
        /// </summary>
        public uint BorderWidthRight { get; set; }

        /// <summary>
        /// Top border width (used only if UseVariableBorders = true)
        /// </summary>
        public uint BorderWidthTop { get; set; }

        /// <summary>
        /// Bottom border width (used only if UseVariableBorders = true)
        /// </summary>
        public uint BorderWidthBottom { get; set; }

        /// <summary>
        /// inside horizontal Border width for table (used only if UseVariableBorders = true)
        /// </summary>
        public uint BorderWidthInsideHorizontal { get; set; }

        /// <summary>
        /// inside vertical border width for table (used only if UseVariableBorders = true)
        /// </summary>
        public uint BorderWidthInsideVertical { get; set; }

        /// <summary>
        /// Border Color in hex value (RRGGBB format)
        /// </summary>
        public string BorderColor { get; set; }

        /// <summary>
        /// Border Color in hex value (RRGGBB format)
        /// </summary>
        public string BorderTopColor { get; set; }

        /// <summary>
        /// Border Color in hex value (RRGGBB format)
        /// </summary>
        public string BorderBottomColor { get; set; }

        /// <summary>
        /// Border Color in hex value (RRGGBB format)
        /// </summary>
        public string BorderLeftColor { get; set; }

        /// <summary>
        /// Border Color in hex value (RRGGBB format)
        /// </summary>
        public string BorderRightColor { get; set; }

        /// <summary>
        /// Default constructor
        /// </summary>
        public BorderModel()
        {
            BorderColor = "000000";
            BorderTopColor = "000000";
            BorderBottomColor = "000000";
            BorderLeftColor = "000000";
            BorderRightColor = "000000";
            BorderWidth = 1;
            BorderWidthTop = 1;
            BorderWidthRight = 1;
            BorderWidthBottom = 1;
            BorderWidthLeft = 1;
        }
    }
}
