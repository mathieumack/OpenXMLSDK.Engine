using OpenXMLSDK.Engine.Word.ReportEngine.Models.Attributes;
using OpenXMLSDK.Engine.Word.Tables.Models;

namespace OpenXMLSDK.Engine.Word.ReportEngine.Models
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
