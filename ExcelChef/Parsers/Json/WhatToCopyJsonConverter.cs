using ExcelChef.Instructions;
using Newtonsoft.Json;
using System;

namespace ExcelChef.Parsers.Json
{
    public class WhatToCopyJsonConverter : JsonConverter
    {
        public override bool CanConvert(Type objectType)
        {
            return typeof(CopyInstruction.WhatToCopy).Equals(objectType);
        }

        public override object ReadJson(JsonReader reader, Type objectType, object existingValue, JsonSerializer serializer)
        {
            string name = serializer.Deserialize<string>(reader);
            return CopyInstruction.WhatToCopy.FromName(name);
        }

        public override void WriteJson(JsonWriter writer, object value, JsonSerializer serializer)
        {
            throw new NotImplementedException();
        }
    }
}
