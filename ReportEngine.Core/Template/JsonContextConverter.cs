using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace ReportEngine.Core.Template
{
    /// <summary>
    /// Json converter
    /// </summary>
    public class JsonContextConverter : JsonConverter
    {
        private readonly IEnumerable<Type> managedTypes;

        /// <summary>
        /// Constructor
        /// </summary>
        public JsonContextConverter()
        {
            managedTypes = typeof(BaseModel).GetTypeInfo().Assembly.GetTypes().Where(t => typeof(BaseModel).IsAssignableFrom(t));
        }

        /// <summary>
        /// Indicate if the converter can be used
        /// </summary>
        /// <param name="objectType">object to unserialize</param>
        /// <returns></returns>
        public override bool CanConvert(Type objectType)
        {
            return objectType == typeof(BaseModel) || objectType == typeof(BaseElement);
        }

        /// <summary>
        /// Read the Json
        /// </summary>
        /// <param name="reader">JsonReader</param>
        /// <param name="objectType">Type</param>
        /// <param name="existingValue">Object to unserialize</param>
        /// <param name="serializer">JsonSerializer</param>
        /// <returns></returns>
        public override object ReadJson(JsonReader reader, Type objectType, object existingValue, JsonSerializer serializer)
        {
            // Load JObject from stream 
            JObject jObject = JObject.Load(reader);

            if (jObject["TypeName"] != null)
            {
                var typeName = jObject["TypeName"].Value<string>();
                if (managedTypes.Any(e => e.Name == typeName))
                    return jObject.ToObject(managedTypes.FirstOrDefault(e => e.Name == typeName), serializer);

                return null;
            }
            else
            {
                throw new Exception("file format exception (Type missing) : " + jObject.First);
            }
        }

        /// <summary>
        /// Write the Json
        /// </summary>
        /// <param name="writer">JsonWriter</param>
        /// <param name="value">Object to serialize</param>
        /// <param name="serializer">JsonSerializer</param>
        public override void WriteJson(JsonWriter writer, object value, JsonSerializer serializer)
        {
            throw new NotImplementedException();
        }
    }
}
