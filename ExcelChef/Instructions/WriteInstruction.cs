using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;

namespace ExcelChef.Instructions
{
    /// <summary>
    /// Writes a value to a cell in a worksheet.
    /// </summary>
    public class WriteInstruction : IInstruction
    {
        /// <summary>
        /// The address of the cell or range the values should be written to.
        /// </summary>
        public string Dst { get; set; }

        /// <summary>
        /// The name or position of the sheet the value should be written to. Defaults to 1.
        /// </summary>
        public object DstSheet { get; set; } = 1L;

        /// <summary>
        /// The values to be written. In row-major order.
        /// </summary>
        public IEnumerable<object> Values { get; set; }

        void IInstruction.Execute(IWorkbook workbook)
        {
            if (string.IsNullOrWhiteSpace(Dst)) throw new Exception($"Destination must be specified");

            ISheet sheet = InstructionUtils.GetSheet(workbook, DstSheet);
            IReadOnlyCollection<ICell> range = InstructionUtils.GetRange(sheet, Dst).ToList();
            IReadOnlyCollection<object> values = Values.ToList();
            if (range.Count != values.Count) throw new Exception($"Destination range must have the same size as the values");
            foreach (var pair in range.Zip(values, (Cell, Value) => new { Cell, Value }))
            {
                InstructionUtils.SetCellValue(pair.Cell, pair.Value);
            }
        }
    }
}
