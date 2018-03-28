﻿using NPOI.SS.UserModel;
using NPOI.SS.Util;
using System;

namespace ExcelChef.Instructions
{
    /// <summary>
    /// Writes a value to a cell in a worksheet.
    /// </summary>
    public class WriteInstruction : IInstruction
    {
        /// <summary>
        /// The address of the cell the value should be written to.
        /// </summary>
        public string Cell { get; set; }

        /// <summary>
        /// The name or position of the worksheet the value should be written to. Defaults to 1.
        /// </summary>
        public object Worksheet { get; set; } = 1L;

        /// <summary>
        /// The value to be written.
        /// </summary>
        public object Value { get; set; }

        void IInstruction.Execute(IWorkbook workbook)
        {
            if (string.IsNullOrWhiteSpace(Cell)) throw new Exception($"{nameof(Cell)} must be specified");

            ISheet worksheet = GetWorksheet(workbook);
            CellReference cellRef = new CellReference(Cell);
            IRow row = worksheet.GetRow(cellRef.Row) ?? worksheet.CreateRow(cellRef.Row);
            ICell cell = row.GetCell(cellRef.Col) ?? row.CreateCell(cellRef.Col);
            ICellStyle style = cell.CellStyle;
            WriteValue(cell);
            cell.CellStyle = style;
            worksheet.ForceFormulaRecalculation = true;
        }

        private ISheet GetWorksheet(IWorkbook workbook)
        {
            switch (Worksheet)
            {
                case long position: return workbook.GetSheetAt((int)position - 1);
                case int position: return workbook.GetSheetAt(position - 1);
                case string name: return workbook.GetSheet(name);
                default: throw new Exception($"{nameof(Worksheet)} must be a name or a position");
            }
        }

        private void WriteValue(ICell cell)
        {
            switch (Value)
            {
                case string text: cell.SetCellValue(text); break;
                case long number: cell.SetCellValue(number); break;
                case double number: cell.SetCellValue(number); break;
                case bool boolean: cell.SetCellValue(boolean); break;
                default: throw new Exception($"{nameof(Value)} must be a string, number or boolean");
            }
        }
    }
}
