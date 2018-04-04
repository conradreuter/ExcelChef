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
                .ToDictionary(TypeFromKind);
        }

        private static string TypeFromKind(Type type)
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
        /// Get a sheet from a workbook. Defaults to the first sheet.
        /// </summary>
        public static ISheet GetSheet(IWorkbook workbook, object sheet)
        {
            switch (sheet)
            {
                case long position: return workbook.GetSheetAt((int)position - 1);
                case int position: return workbook.GetSheetAt(position - 1);
                case string name: return workbook.GetSheet(name);
                default: throw new Exception($@"Expected name or a position for sheet, got ""{sheet}""");
            }
        }

        /// <summary>
        /// Get a cell from a sheet.
        /// </summary>
        public static ICell GetCell(ISheet sheet, string cell)
        {
            CellReference address = new CellReference(cell);
            IRow row = sheet.GetRow(address.Row) ?? sheet.CreateRow(address.Row);
            return row.GetCell(address.Col) ?? row.CreateCell(address.Col);
        }

        /// <summary>
        /// Get a range from a sheet.
        /// </summary>
        public static IEnumerable<ICell> GetRange(ISheet sheet, string range)
        {
            CellRangeAddress address = CellRangeAddress.ValueOf(range);
            for (int rowIdx = address.FirstRow; rowIdx <= address.LastRow; ++rowIdx)
            {
                IRow row = sheet.GetRow(rowIdx) ?? sheet.CreateRow(rowIdx);
                for (int colIdx = address.FirstColumn; colIdx <= address.LastColumn; ++colIdx)
                {
                    yield return row.GetCell(colIdx) ?? row.CreateCell(colIdx);
                }
            }
        }

        /// <summary>
        /// Set the value of a cell.
        /// </summary>
        public static void SetCellValue(ICell cell, object value)
        {
            switch (value)
            {
                case null: cell.SetCellType(CellType.Blank); break;
                case string text: cell.SetCellValue(text); break;
                case long number: cell.SetCellValue(number); break;
                case double number: cell.SetCellValue(number); break;
                case bool boolean: cell.SetCellValue(boolean); break;
                default: throw new Exception($@"Value must be a string, number or boolean, got ""{value}""");
            }
        }

        /// <summary>
        /// Get the value of a cell.
        /// </summary>
        public static object GetCellValue(ICell cell)
        {
            if (cell.CellType == CellType.Formula)
            {
                return GetCellValueAssumingType(cell, cell.CachedFormulaResultType);
            }
            return GetCellValueAssumingType(cell, cell.CellType);
        }

        private static object GetCellValueAssumingType(ICell cell, CellType cellType)
        {
            switch (cellType)
            {
                case CellType.Boolean: return cell.BooleanCellValue;
                case CellType.Error: return cell.ErrorCellValue;
                case CellType.Numeric: return cell.NumericCellValue;
                case CellType.String: return cell.StringCellValue;
                default: return null;
            }
        }
    }
}
