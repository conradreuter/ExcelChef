using NPOI.SS.UserModel;
using System;

namespace ExcelChef.Instructions
{
    /// <summary>
    /// Invalidates cached formula results.
    /// </summary>
    public class InvalidateInstruction : IInstruction
    {
        /// <summary>
        /// The address of the cell or range that should be invalidated.
        /// </summary>
        public string Dst { get; set; }

        /// <summary>
        /// The name or position of the sheet where cells should be invalidated. Defaults to 1.
        /// </summary>
        public object DstSheet { get; set; } = 1L;

        void IInstruction.Execute(IWorkbook workbook)
        {
            if (string.IsNullOrWhiteSpace(Dst)) throw new Exception($"Destination must be specified");

            ISheet sheet = InstructionUtils.GetSheet(workbook, DstSheet);
            foreach (ICell cell in InstructionUtils.GetRange(sheet, Dst, excludeEmpty: true))
            {
                if (cell.CellType == CellType.Formula)
                {
                    string formula = cell.CellFormula;
                    cell.SetCellType(CellType.Blank);
                    cell.SetCellFormula(formula);
                }
            }
        }
    }
}
