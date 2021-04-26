namespace OpenXMLSDK.Engine.Word.ReportEngine.Models
{
    /// <summary>
    /// Uniform Grid
    /// </summary>
    public class UniformGrid : Table
    {
        /// <summary>
        /// Model of each cell
        /// </summary>
        public Cell CellModel { get; set; }

        /// <summary>
        /// Indicate if rows can be splited in multiple pages
        /// </summary>
        public bool CantSplitRows { get; set; }

        /// <summary>
        /// Key indicating the number of column for spliting the grid 
        /// </summary>
        public string ColumnNumberKey { get; set; }

        /// <summary>
        /// Key indicating if there are column Headers
        /// </summary>
        public string AreColumnHeadersKey { get; set; }

        /// <summary>
        /// Key indicating if there are row Headers
        /// </summary>
        public string AreRowHeadersKey { get; set; }

        /// <summary>
        /// Headers background color in hex value (RRGGBB format)
        /// </summary>
        public string HeadersColor { get; set; }

        /// <summary>
        /// Constructor
        /// </summary>
        public UniformGrid()
            : base(typeof(UniformGrid).Name)
        {
        }
    }
}
