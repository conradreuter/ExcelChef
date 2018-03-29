using ExcelChef.Instructions;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;

namespace ExcelChef.Parsers.Json
{
    public class InstructionJsonConverter : JsonConverter
    {
        public override bool CanConvert(Type objectType)
        {
            return typeof(IInstruction).Equals(objectType);
        }

        public override object ReadJson(JsonReader reader, Type objectType, object existingValue, JsonSerializer serializer)
        {
            JObject jsonObject = serializer.Deserialize<JObject>(reader);
            string kind = jsonObject.Property("kind").Value.ToObject<string>();
            Type type = InstructionUtils.GetType(kind);
            return jsonObject.ToObject(type, serializer);
        }

        public override void WriteJson(JsonWriter writer, object value, JsonSerializer serializer)
        {
            throw new NotImplementedException();
        }
    }
}
