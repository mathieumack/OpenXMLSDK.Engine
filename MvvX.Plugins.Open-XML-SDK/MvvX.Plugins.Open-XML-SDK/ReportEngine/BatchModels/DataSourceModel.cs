using System.Collections.Generic;

namespace MvvX.Plugins.OpenXMLSDK.ReportEngine.BatchModels
{
    public class DataSourceModel : BaseModel
    {
        public List<ContextModel> Items { get; set; }

        public DataSourceModel()
            : this(new List<ContextModel>())
        { }

        public DataSourceModel(List<ContextModel> items)
            : base(typeof(DataSourceModel).Name)
        {
            Items = items;
        }
    }
}
