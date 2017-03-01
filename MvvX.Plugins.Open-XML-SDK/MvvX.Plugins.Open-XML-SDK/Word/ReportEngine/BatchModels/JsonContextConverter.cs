using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using MvvmCross.Platform;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace MvvX.Plugins.OpenXMLSDK.Word.ReportEngine.BatchModels
{
    public class JsonContextConverter : JsonConverter
    {
        private readonly IEnumerable<Type> managedTypes;

        public JsonContextConverter()
        {
            managedTypes = typeof(BaseModel).GetTypeInfo().Assembly.GetTypes().Where(t => typeof(BaseModel).IsAssignableFrom(t));
        }

        public override bool CanConvert(Type objectType)
        {
            return objectType == typeof(BaseModel);
        }

        public override object ReadJson(JsonReader reader, Type objectType, object existingValue, JsonSerializer serializer)
        {
            // Load JObject from stream 
            JObject jObject = JObject.Load(reader);

            if (jObject["Type"] != null)
            {
                var typeName = jObject["Type"].Value<string>();
                if (managedTypes.Any(e => e.Name == typeName))
                    return jObject.ToObject(managedTypes.FirstOrDefault(e => e.Name == typeName), serializer);

                return null;
            }
            else
            {
                throw new Exception("file format exception (Type missing) : " + jObject.First);
            }
        }

        public override void WriteJson(JsonWriter writer, object value, JsonSerializer serializer)
        {
            throw new NotImplementedException();
        }
    }
}
