namespace MvvX.Open_XML_SDK.Core.Word.Tables.Models
{
    public class TableCellPropertiesModel
    {
        public string Width { get; set;}

        public TableWidthUnitValues WidthUnit { get; set; }

        public Shading Shading { get; set; }

        public TableVerticalAlignmentValues TableVerticalAlignementValues { get; set; }

        public JustificationValues? Justification { get; set; }

        public TextDirectionValues? TextDirectionValues { get; set; }

        public int? Gridspan { get; set; }

        public bool Fusion { get; set; }

        public bool FusionChild { get; set; }

        public string Height { get; set; }

        /// <summary>
        /// Permet de rendre la cellule suivante solidaire avec celle-ci (En cas de saut de page par exemple). Par défaut à false
        /// </summary>
        public bool ParagraphSolidarity { get; set; }

        /// <summary>
        /// Gestion de la bordure du haut de la cellule. De base à null
        /// </summary>
        public TableBorderModel TopBorder { get; set; }

        /// <summary>
        /// Gestion de la bordure du bas de la cellule. De base à null
        /// </summary>
        public TableBorderModel BottomBorder { get; set; }

        /// <summary>
        /// Gestion de la bordure de gauche de la cellule. De base à null
        /// </summary>
        public TableBorderModel LeftBorder { get; set; }

        /// <summary>
        /// Gestion de la bordure de droite de la cellule. De base à null
        /// </summary>
        public TableBorderModel RightBorder { get; set; }

        public TableCellPropertiesModel()
        {
            TableVerticalAlignementValues = TableVerticalAlignmentValues.Top;
            ParagraphSolidarity = false;
        }

    }
}
