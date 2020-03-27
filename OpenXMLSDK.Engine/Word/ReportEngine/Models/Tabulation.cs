namespace OpenXMLSDK.Engine.Word.ReportEngine.Models
{
    /// <summary>
    /// Model for Tabulation
    /// </summary>
    public class Tabulation : BaseElement
    {
        /// <summary>
        /// Tabulation alignment 
        /// </summary>
        public TabAlignmentValues Alignment { get; set; }

        /// <summary>
        /// Tabulation stop leader character value
        /// </summary>
        public TabStopLeaderCharValues Leader { get; set; } = TabStopLeaderCharValues.none;

        /// <summary>
        /// Tabulation stop position
        /// </summary>
        public int TabStopPosition { get; set; }

        /// <summary>
        /// Content text
        /// </summary>
        public Label Text { get; set; } = new Label();

        /// <summary>
        /// Constructor
        /// </summary>
        public Tabulation()
            : base(typeof(Tabulation).Name)
        {
        }
    }
}
