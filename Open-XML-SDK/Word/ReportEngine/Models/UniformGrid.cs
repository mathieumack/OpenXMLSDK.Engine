using MvvX.Plugins.OpenXMLSDK.Word.ReportEngine.Models.Attributes;
using MvvX.Plugins.OpenXMLSDK.Word.Tables.Models;

namespace MvvX.Plugins.OpenXMLSDK.Word.ReportEngine.Models
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
