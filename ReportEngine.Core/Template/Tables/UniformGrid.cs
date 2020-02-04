namespace ReportEngine.Core.Template.Tables
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
        /// Constructor
        /// </summary>
        public UniformGrid()
            : base(typeof(UniformGrid).Name)
        {
        }
    }
}
