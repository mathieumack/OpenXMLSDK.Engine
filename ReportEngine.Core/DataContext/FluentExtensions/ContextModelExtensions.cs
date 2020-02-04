using System;
using System.Linq;

namespace ReportEngine.Core.DataContext.FluentExtensions
{
    public static class ContextModelExtensions
    {
        /// <summary>
        /// Add string model
        /// </summary>
        /// <param name="context"></param>
        /// <param name="key"></param>
        /// <param name="value"></param>
        /// <returns></returns>
        public static ContextModel AddString(this ContextModel context, string key, string value)
        {
            var element = new StringModel(value);
            context.AddItem(key, element);
            return context;
        }

        /// <summary>
        /// Add double model
        /// </summary>
        /// <param name="context"></param>
        /// <param name="key"></param>
        /// <param name="value"></param>
        /// <param name="renderPattern"></param>
        /// <returns></returns>
        public static ContextModel AddDouble(this ContextModel context, string key, double value, string renderPattern)
        {
            var element = new DoubleModel(value, renderPattern);
            context.AddItem(key, element);
            return context;
        }

        /// <summary>
        /// Add double model
        /// </summary>
        /// <param name="context"></param>
        /// <param name="key"></param>
        /// <param name="value"></param>
        /// <returns></returns>
        public static ContextModel AddBoolean(this ContextModel context, string key, bool value)
        {
            var element = new BooleanModel(value);
            context.AddItem(key, element);
            return context;
        }

        /// <summary>
        /// Add Date time model
        /// </summary>
        /// <param name="context"></param>
        /// <param name="key"></param>
        /// <param name="value"></param>
        /// <param name="renderPattern"></param>
        /// <returns></returns>
        public static ContextModel AddDateTime(this ContextModel context, string key, DateTime value, string renderPattern)
        {
            var element = new DateTimeModel(value, renderPattern);
            context.AddItem(key, element);
            return context;
        }

        /// <summary>
        /// Add Date time model
        /// </summary>
        /// <param name="context"></param>
        /// <param name="key"></param>
        /// <param name="value"></param>
        /// <returns></returns>
        public static ContextModel AddByteContent(this ContextModel context, string key, byte[] value)
        {
            var element = new ByteContentModel(value);
            context.AddItem(key, element);
            return context;
        }

        /// <summary>
        /// Add Date time model
        /// </summary>
        /// <param name="context"></param>
        /// <param name="key"></param>
        /// <param name="value"></param>
        /// <returns></returns>
        public static ContextModel AddBase64Content(this ContextModel context, string key, string value)
        {
            var element = new Base64ContentModel(value);
            context.AddItem(key, element);
            return context;
        }

        /// <summary>
        /// Add a list of elements as a DataSource
        /// </summary>
        /// <param name="context"></param>
        /// <param name="key"></param>
        /// <param name="elements"></param>
        /// <returns></returns>
        public static ContextModel AddCollection(this ContextModel context, string key, params ContextModel[] elements)
        {
            var element = new DataSourceModel()
            {
                Items = elements.ToList()
            };
            context.AddItem(key, element);
            return context;
        }
    }
}
