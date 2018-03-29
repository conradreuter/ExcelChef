using NPOI.SS.UserModel;
using NPOI.SS.Util;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace ExcelChef.Instructions
{
    /// <summary>
    /// Helper methods for all instructions.
    /// </summary>
    public static class InstructionUtils
    {
        private static readonly Regex _kindRegex = new Regex("^(.*)Instruction$");
        private static readonly IReadOnlyDictionary<string, Type> _typesByKind;

        static InstructionUtils()
        {
            _typesByKind =
                typeof(IInstruction).Assembly.GetTypes()
                .Where(typeof(IInstruction).IsAssignableFrom)
                .Where(t => !t.IsAbstract)
                .ToDictionary(GetKind);
        }

        private static string GetKind(Type type)
        {
            Match match = _kindRegex.Match(type.Name);
            if (!match.Success)
            {
                throw new Exception($@"Instruction type ""{type.Name}"" does not match the expected pattern");
            }
            return match.Groups[1].Value.ToLowerInvariant();
        }

        /// <summary>
        /// Get the instruction type from a kind.
        /// </summary>
        public static Type GetType(string kind)
        {
            kind = kind.ToLowerInvariant();
            if (!_typesByKind.TryGetValue(kind, out Type type))
            {
                throw new Exception($@"Unknown instruction kind ""{kind}""");
            }
            return type;
        }

        /// <summary>
        /// Get a worksheet from a workbook.
        /// </summary>
        public static ISheet GetWorksheet(IWorkbook workbook, object worksheet)
        {
            switch (worksheet)
            {
                case long position: return workbook.GetSheetAt((int)position - 1);
                case int position: return workbook.GetSheetAt(position - 1);
                case string name: return workbook.GetSheet(name);
                default: throw new Exception($"{nameof(worksheet)} must be a name or a position");
            }
        }

        /// <summary>
        /// Get a cell from a worksheet.
        /// </summary>
        public static ICell GetCell(ISheet worksheet, string cell)
        {
            CellReference cellReference = new CellReference(cell);
            IRow row = worksheet.GetRow(cellReference.Row) ?? worksheet.CreateRow(cellReference.Row);
            return row.GetCell(cellReference.Col) ?? row.CreateCell(cellReference.Col);
        }
    }
}
