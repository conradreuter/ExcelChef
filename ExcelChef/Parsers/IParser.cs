using System.Collections.Generic;
using System.IO;

namespace ExcelChef.Parsers
{
    /// <summary>
    /// Parses the programs input.
    /// </summary>
    public interface IParser
    {
        /// <summary>
        /// Parse the program input and provide instructions.
        /// </summary>
        IEnumerable<IInstruction> Parse(TextReader input);
    }
}
