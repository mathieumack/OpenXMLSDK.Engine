using System;
using System.Linq;

namespace OpenXMLSDK.Engine.ReportEngine.DataContext.FluentExtensions
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
        /// Add a file link model (image path or any file reference used by the report engine)
        /// </summary>
        /// <param name="context"></param>
        /// <param name="key"></param>
        /// <param name="filePath">Absolute or relative path to the file</param>
        /// <returns></returns>
        public static ContextModel AddFileLink(this ContextModel context, string key, string filePath)
        {
            var element = new FileLinkModel(filePath);
            context.AddItem(key, element);
            return context;
        }

        /// <summary>
        /// Add a substitutable string model, allowing composite formatted strings built from other context values
        /// </summary>
        /// <param name="context"></param>
        /// <param name="key"></param>
        /// <param name="renderPattern">Format pattern (e.g. "{0} of {1}")</param>
        /// <param name="dataSource">Context whose values are injected into the pattern in order</param>
        /// <returns></returns>
        public static ContextModel AddSubstitutableString(this ContextModel context, string key, string renderPattern, ContextModel dataSource)
        {
            var element = new SubstitutableStringModel(renderPattern, dataSource);
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
