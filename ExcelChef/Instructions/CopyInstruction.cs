using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;

namespace ExcelChef.Instructions
{
    /// <summary>
    /// Copies aspects of a cell or range to a cell or equally sized range.
    /// </summary>
    public class CopyInstruction : IInstruction
    {
        /// <summary>
        /// The address of the cell or range to be copied to.
        /// </summary>
        public string Dst { get; set; }

        /// <summary>
        /// The name or position of the sheet to be copied to. Defaults to 1.
        /// </summary>
        public object DstSheet { get; set; } = 1L;

        /// <summary>
        /// The address of the cell or range to be copied from.
        /// </summary>
        public string Src { get; set; }

        /// <summary>
        /// The name or position of the sheet to be copied from. Defaults to 1.
        /// </summary>
        public object SrcSheet { get; set; } = 1L;

        /// <summary>
        /// What to be copied. Defaults to formulas and styles.
        /// </summary>
        public ISet<WhatToCopy> What { get; set; }

        void IInstruction.Execute(IWorkbook workbook)
        {
            if (Src == null) throw new Exception($"The source range must be specified");
            if (Dst == null) throw new Exception($"The destination range must be specified");
            if (What == null) What = new HashSet<WhatToCopy> { WhatToCopy.Formulas, WhatToCopy.Styles };

            ISheet srcSheet = InstructionUtils.GetSheet(workbook, SrcSheet);
            ISheet dstSheet = InstructionUtils.GetSheet(workbook, DstSheet);
            IReadOnlyCollection<ICell> srcRange = InstructionUtils.GetRange(srcSheet, Src).ToList();
            IReadOnlyCollection<ICell> dstRange = InstructionUtils.GetRange(dstSheet, Dst).ToList();
            if (srcRange.Count == 1) srcRange = Enumerable.Repeat(srcRange.First(), dstRange.Count).ToList();
            if (srcRange.Count != dstRange.Count) throw new Exception($"Source and destination range must have the same dimensions");
            foreach (var pair in srcRange.Zip(dstRange, (Src, Dst) => new { Src, Dst }))
            {
                foreach (WhatToCopy whatToCopy in What)
                {
                    whatToCopy.Execute(pair.Src, pair.Dst);
                }
            }
        }

        /// <summary>
        /// Things that can be copied.
        /// </summary>
        public class WhatToCopy
        {
            private static readonly IReadOnlyDictionary<string, WhatToCopy> _instancesByName;

            static WhatToCopy()
            {
                _instancesByName =
                    typeof(WhatToCopy).GetFields(BindingFlags.Public | BindingFlags.Static)
                    .Where(f => typeof(WhatToCopy).Equals(f.FieldType))
                    .ToDictionary(f => f.Name.ToLowerInvariant(), f => (WhatToCopy)f.GetValue(null));
            }

            /// <summary>
            /// Copy formulas (or values if there is no formula).
            /// </summary>
            public static readonly WhatToCopy Formulas = new WhatToCopy(CopyFormula);

            private static void CopyFormula(ICell srcCell, ICell dstCell)
            {
                CopyValue(srcCell, dstCell);
                if (srcCell.CellType == CellType.Formula)
                {
                    dstCell.SetCellFormula(srcCell.CellFormula);
                }
            }

            /// <summary>
            /// Copy styles, i.e. number formats, colors. conditional formatting, etc.
            /// </summary>
            public static readonly WhatToCopy Styles = new WhatToCopy(CopyStyle);

            private static void CopyStyle(ICell srcCell, ICell dstCell)
            {
                dstCell.CellStyle = srcCell.CellStyle;
            }

            /// <summary>
            /// Copy values.
            /// </summary>
            public static readonly WhatToCopy Values = new WhatToCopy(CopyValue);

            private static void CopyValue(ICell srcCell, ICell dstCell)
            {
                object value = InstructionUtils.GetCellValue(srcCell);
                InstructionUtils.SetCellValue(dstCell, value);
            }

            /// <summary>
            /// Get the instance from a name.
            /// </summary>
            public static WhatToCopy FromName(string name)
            {
                name = name.ToLowerInvariant();
                if (!_instancesByName.TryGetValue(name, out WhatToCopy instance))
                {
                    throw new Exception($@"Unknown copy instruction ""{name}""");
                }
                return instance;
            }

            private WhatToCopy(Action<ICell, ICell> execute)
            {
                Execute = execute;
            }

            /// <summary>
            /// Execute the copy operation.
            /// </summary>
            public Action<ICell, ICell> Execute { get; }
        }
    }
}
