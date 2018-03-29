using NPOI.SS.UserModel;
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
        /// The name or position of the sheet the value should be written to. Defaults to 1.
        /// </summary>
        public object Sheet { get; set; } = 1L;

        /// <summary>
        /// The value to be written.
        /// </summary>
        public object Value { get; set; }

        void IInstruction.Execute(IWorkbook workbook)
        {
            if (string.IsNullOrWhiteSpace(Cell)) throw new Exception($"Cell must be specified");

            ISheet sheet = InstructionUtils.GetSheet(workbook, Sheet);
            ICell cell = InstructionUtils.GetCell(sheet, Cell);
            InstructionUtils.SetCellValue(cell, Value);
        }
    }
}
