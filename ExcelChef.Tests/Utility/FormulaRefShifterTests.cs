using FluentAssertions;
using NUnit.Framework;
using static ExcelChef.Utility.FormulaRefShifter;

namespace ExcelChef.Utility
{
    [TestFixture]
    public class FormulaRefShifterTests
    {
        [Test]
        public void ShiftFormulaRefs_ShiftsTheRowInReferences()
        {
            // assert
            ShiftFormulaRefs("X42", 1, 0).Should().Be("X43");
        }

        [Test]
        public void ShiftFormulaRefs_ShiftsTheRowInRangeReferences()
        {
            // assert
            ShiftFormulaRefs("X42:X1337", 1, 0).Should().Be("X43:X1338");
        }

        [Test]
        public void ShiftFormulaRefs_LeavesAnchoredRowReferencesIntact()
        {
            // assert
            ShiftFormulaRefs("X$42", 1, 0).Should().Be("X$42");
        }

        [Test]
        public void ShiftFormulaRefs_ShiftsTheColumnInReferences()
        {
            // assert
            ShiftFormulaRefs("X42", 0, 1).Should().Be("Y42");
        }

        [Test]
        public void ShiftFormulaRefs_ShiftsTheColumnInRangeReferences()
        {
            // assert
            ShiftFormulaRefs("X42:AA42", 0, 1).Should().Be("Y42:AB42");
        }

        [Test]
        public void ShiftFormulaRefs_LeavesAnchoredColumnReferencesIntact()
        {
            // assert
            ShiftFormulaRefs("$X42", 0, 1).Should().Be("$X42");
        }

        [Test]
        public void ShiftFormulaRefs_LeavesFunctionsIntact()
        {
            // assert
            ShiftFormulaRefs("ABS(X42)", 1, 1).Should().Be("ABS(Y43)");
        }
    }
}
