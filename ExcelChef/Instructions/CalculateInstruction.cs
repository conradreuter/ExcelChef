using NPOI.SS.UserModel;

namespace ExcelChef.Instructions
{
    /// <summary>
    /// Calculates all sheets in a workbook.
    /// </summary>
    public class CalculateInstruction : IInstruction
    {
        void IInstruction.Execute(IWorkbook workbook)
        {
            workbook.GetCreationHelper().CreateFormulaEvaluator().EvaluateAll();
        }
    }
}
