using Newtonsoft.Json;
using Newtonsoft.Json.Serialization;
using System.Collections.Generic;
using System.IO;

namespace ExcelChef.Parsers.Json
{
    /// <summary>
    /// Parses the program input as JSON.
    /// </summary>
    public class JsonParser : IParser
    {
        private readonly JsonSerializer _jsonSerializer = new JsonSerializer
        {
            ContractResolver = new CamelCasePropertyNamesContractResolver(),
            Converters =
            {
                new InstructionJsonConverter(),
                new WhatToCopyJsonConverter(),
            },
        };

        IEnumerable<IInstruction> IParser.Parse(TextReader input)
        {
            JsonReader jsonReader = new JsonTextReader(input);
            return _jsonSerializer.Deserialize<IEnumerable<IInstruction>>(jsonReader);
        }
    }
}
