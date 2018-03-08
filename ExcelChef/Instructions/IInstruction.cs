using NPOI.SS.UserModel;

namespace ExcelChef
{
    /// <summary>
    /// An instruction provided by the program input.
    /// </summary>
    public interface IInstruction
    {
        /// <summary>
        /// Execute the instruction on a workbook.
        /// </summary>
        void Execute(IWorkbook workbook);
    }
}
