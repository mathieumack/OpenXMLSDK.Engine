using System.Collections.Generic;

namespace ReportEngine.Core.Template
{
    public class ForEach : BaseElement
    {
        /// <summary>
        /// Define the prefix that will be used for automatically added items
        /// IsFirstItem, ...
        /// </summary>
        public string AutoContextAddItemsPrefix { get; set; }

        /// <summary>
        /// Describe the template of each final item
        /// Warning : If the ForEach control is under a Cell, the ItemTemplate must be a Paragraph !
        /// </summary>
        public List<BaseElement> ItemTemplate { get; set; }

        /// <summary>
        /// Items that will be repeated
        /// </summary>
        public string DataSourceKey { get; set; }

        public ForEach()
            : base(typeof(ForEach).Name)
        {
        }
    }
}