using OpenXMLSDK.Word.ReportEngine.Models.Attributes;
using OpenXMLSDK.Word.Tables.Models;

namespace OpenXMLSDK.Word.ReportEngine.Models
{
    public class UniformGrid : Table
    {
        public Cell CellModel { get; set; }

        /// <summary>
        /// Constructor
        /// </summary>
        public UniformGrid() 
            : base(typeof(UniformGrid).Name)
        {
        }
    }
}
