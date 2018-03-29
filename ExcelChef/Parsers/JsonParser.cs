using ExcelChef.Instructions;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using Newtonsoft.Json.Serialization;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace ExcelChef.Parsers
{
    /// <summary>
    /// Parses the program input as JSON.
    /// </summary>
    public class JsonParser : IParser
    {
        private readonly JsonSerializer _jsonSerializer = new JsonSerializer
        {
            ContractResolver = new CamelCasePropertyNamesContractResolver(),
        };

        IEnumerable<IInstruction> IParser.Parse(TextReader input)
        {
            JsonReader jsonReader = new JsonTextReader(input);
            JEnumerable<JObject> jsonArray = _jsonSerializer.Deserialize<JEnumerable<JObject>>(jsonReader);
            return jsonArray.Select(Parse);
        }

        private IInstruction Parse(JObject jsonObject)
        {
            string kind = jsonObject.Property("kind").Value.ToObject<string>();
            Type type = InstructionUtils.GetType(kind);
            return (IInstruction)jsonObject.ToObject(type);
        }
    }
}
