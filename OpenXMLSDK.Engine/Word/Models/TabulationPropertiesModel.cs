using OpenXMLSDK.Engine.Word.ReportEngine.Models;

namespace OpenXMLSDK.Engine.Word.Models
{
    /// <summary>
    /// Model for Tabulation properties
    /// </summary>
    public class TabulationPropertiesModel
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
    }
}
