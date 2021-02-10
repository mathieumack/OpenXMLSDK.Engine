namespace OpenXMLSDK.Engine.Word.ReportEngine.Models
{
    /// <summary>
    /// Model for a column
    /// </summary>
    public class Column : BaseElement
    {
        /// <summary>
        /// Are columns width the same 
        /// </summary>
        public bool EqualWidth { get; set; }

        /// <summary>
        /// Number of column 
        /// </summary>
        public ColumnCountValues Number { get; set; }

        /// <summary>
        /// Constructor
        /// </summary>
        public Column()
            : base(typeof(Column).Name)
        {
            // EqualWidth set to true by default
            EqualWidth = true;
        }
    }
}
