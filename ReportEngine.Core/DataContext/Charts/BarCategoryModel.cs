using System;

namespace ReportEngine.Core.DataContext.Charts
{
    [Obsolete("Please use CategoryModel instead")]
    public class BarCategoryModel
    {
        /// <summary>
        /// Name of the category
        /// </summary>
        public string Name { get; set; }

        /// <summary>
        /// Category color
        /// </summary>
        public string Color { get; set; }
    }
}
