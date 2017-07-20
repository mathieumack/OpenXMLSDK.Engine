using System.Collections.Generic;

namespace MvvX.Plugins.OpenXMLSDK.Word.Models.Charts
{
    public class ChartModel
    {
        /// <summary>
        /// List of categories
        /// </summary>
        public IList<ChartCategorie> Categories { get; set; }

        /// <summary>
        /// Values series
        /// </summary>
        public IList<ChartSerie> Series { get; set; }

        /// <summary>
        /// Indicate if the legend must be included
        /// </summary>
        public bool ShowLegend { get; set; }

        /// <summary>
        /// Permet de mettre en place une Font sur la légende du graphique
        /// </summary>
        public string FontFamilyLegend { get; set; }

        /// <summary>
        /// Titre du graphique
        /// </summary>
        public string Title { get; set; }

        /// <summary>
        /// Valeur max en pixel de la largeur du graphique. (Peut être à null)
        /// </summary>
        public long? MaxWidth { get; set; }

        /// <summary>
        /// Valeur max en pixel de la hauteur du graphique. (Peut être à null)
        /// </summary>
        public long? MaxHeight { get; set; }

        /// <summary>
        /// Booléan indiquant si il faut affiché ou non les valeurs en chiffre
        /// </summary>
        public bool ShowDataLabel { get; set; }

        /// <summary>
        /// Couleur des valeurs en chiffre. Seulement si ShowDataLabel est à true (Peut être à null)
        /// </summary>
        public string DataLabelColor { get; set; }

        /// <summary>
        /// Booléan indiquant si il faut afficher ou non le titre du graphique
        /// </summary>
        public bool ShowTitle { get; set; }

        /// <summary>
        /// Booléan indiquant si il faut afficher ou non la bordure du graphique
        /// </summary>
        public bool BoderColor { get; set; }

        /// <summary>
        /// Booléan indiquant s'il faut supprimer ou non la/les catégorie(s) sur l'axe Y 
        /// </summary>
        public bool DeleteAxeCategorie { get; set; }

        /// <summary>
        /// Booléan indiquant s'il faut supprimer ou non la/les valeur(s) sur l'axe X 
        /// </summary>
        public bool DeleteAxeValue { get; set; }

        /// <summary>
        /// Booléan indiquant s'il faut afficher ou non les MajorGridlines (Ligne vertical au milieu du graphique)
        /// </summary>
        public bool ShowMajorGridlines { get; set; }

        /// <summary>
        /// Valeur de l'espace entre les lignes dans catégories dans un barChart. (Peut être à null)
        /// </summary>
        public int? SpaceBetweenLineCategorie { get; set; }
    }
}
