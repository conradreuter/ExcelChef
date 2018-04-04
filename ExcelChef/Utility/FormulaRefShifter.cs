using NPOI.SS.Formula;
using NPOI.SS.Formula.PTG;

namespace ExcelChef.Utility
{
    /// <summary>
    /// Shifts the relative references in formulas.
    /// </summary>
    public class FormulaRefShifter
    {
        private readonly int _colOffset;
        private readonly int _rowOffset;

        private FormulaRefShifter(int rowOffset, int colOffset)
        {
            _colOffset = colOffset;
            _rowOffset = rowOffset;
        }

        /// <summary>
        /// Shift the relative references in the given formula.
        /// </summary>
        public static string ShiftFormulaRefs(string formula, int rowOffset, int colOffset)
        {
            return new FormulaRefShifter(rowOffset, colOffset).ShiftFormulaRefs(formula);
        }

        private string ShiftFormulaRefs(string formula)
        {
            Ptg[] tokens = FormulaParser.Parse(formula, null);
            foreach (Ptg token in tokens) ShiftTokenRefs(token);
            return FormulaRenderer.ToFormulaString(null, tokens);
        }

        private void ShiftTokenRefs(Ptg token)
        {
            switch (token)
            {
                case RefPtgBase refToken:
                    if (refToken.IsRowRelative) refToken.Row += _rowOffset;
                    if (refToken.IsColRelative) refToken.Column += _colOffset;
                    break;

                case Area2DPtgBase areaToken:
                    if (areaToken.IsFirstRowRelative) areaToken.FirstRow += _rowOffset;
                    if (areaToken.IsFirstColRelative) areaToken.FirstColumn += _colOffset;
                    if (areaToken.IsLastRowRelative) areaToken.LastRow += _rowOffset;
                    if (areaToken.IsLastColRelative) areaToken.LastColumn += _colOffset;
                    break;
            }
        }
    }
}
