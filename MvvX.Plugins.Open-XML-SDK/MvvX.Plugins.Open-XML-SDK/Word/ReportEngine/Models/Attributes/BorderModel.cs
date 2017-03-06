using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MvvX.Plugins.OpenXMLSDK.Word.ReportEngine.Models.Attributes
{
    public class BorderModel
    {
        /// <summary>
        /// Définie les bordures à utiliser : 0 = None, 1 = Top, 2 = Bottom, 4 = Left, 8 = Right
        /// </summary>
        public BorderPositions BorderPositions { get; set; }

        /// <summary>
        /// border width size in eighths of a point 
        /// </summary>
        public uint BorderWidth { get; set; }

        /// <summary>
        /// Indique si les bordures doivent de la même taille ou non
        /// </summary>
        public bool UseVariableBorders { get; set; }

        /// <summary>
        /// Epaisseur de la bordure de gauche
        /// </summary>
        public uint BorderWidthLeft { get; set; }

        /// <summary>
        /// Epaisseur de la bordure de droite
        /// </summary>
        public uint BorderWidthRight { get; set; }

        /// <summary>
        /// Epaisseur de la bordure du haut
        /// </summary>
        public uint BorderWidthTop { get; set; }

        /// <summary>
        /// Epaisseur de la bordure du bas
        /// </summary>
        public uint BorderWidthBottom { get; set; }

        /// <summary>
        /// inside horizontal Border width for table
        /// </summary>
        public uint BorderWidthInsideHorizontal { get; set; }

        /// <summary>
        /// inside vertical border width for table
        /// </summary>
        public uint BorderWidthInsideVertical { get; set; }

        /// <summary>
        /// Couleur des bordures
        /// </summary>
        public string BorderColor { get; set; }

        /// <summary>
        /// Default constructor
        /// </summary>
        public BorderModel()
        {
            BorderColor = "FFFFFF";
            BorderWidth = 1;
            BorderWidthTop = 1;
            BorderWidthRight = 1;
            BorderWidthBottom = 1;
            BorderWidthLeft = 1;
        }
    }
}
