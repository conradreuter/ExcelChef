using FluentAssertions;
using NUnit.Framework;

namespace ExcelChef.IntegrationTests
{
    [TestFixture]
    public class CopyInstructionIntegrationTests : IntegrationTestsBase
    {
        [Test]
        public void DefaultsToUsingTheFirstSheet()
        {
            // act
            Run(@"
                [
                    {
                        ""kind"": ""copy"",
                        ""src"": ""A1"",
                        ""dst"": ""B1"",
                    }
                ]
            ");

            // assert
            _workbook.GetSheetAt(0).GetRow(0).GetCell(1).StringCellValue.Should().Be("normal");
        }

        [Test]
        public void DefaultsToCopyingFormulasAndStyles()
        {
            // act
            Run(@"
                [
                    {
                        ""kind"": ""copy"",
                        ""src"": ""A2"",
                        ""dst"": ""B1"",
                    },
                    {
                        ""kind"": ""copy"",
                        ""src"": ""A4"",
                        ""dst"": ""B2"",
                    }
                ]
            ");

            // assert
            _workbook.GetSheetAt(0).GetRow(0).GetCell(1).StringCellValue.Should().Be("bold");
            _workbook.GetSheetAt(0).GetRow(0).GetCell(1).CellStyle.GetFont(_workbook).IsBold.Should().BeTrue();
            _workbook.GetSheetAt(0).GetRow(1).GetCell(1).CellFormula.Should().Be("RAND()");
        }

        [Test]
        public void CanReferenceTheSourceSheetByPosition()
        {
            // act
            Run(@"
                [
                    {
                        ""kind"": ""copy"",
                        ""srcSheet"": 1,
                        ""src"": ""A1"",
                        ""dst"": ""B1"",
                        ""what"": [""values""]
                    }
                ]
            ");

            // assert
            _workbook.GetSheetAt(0).GetRow(0).GetCell(1).StringCellValue.Should().Be("normal");
        }

        [Test]
        public void CanReferenceTheDestinationSheetByPosition()
        {
            // act
            Run(@"
                [
                    {
                        ""kind"": ""copy"",
                        ""src"": ""A1"",
                        ""dstSheet"": 1,
                        ""dst"": ""B1"",
                        ""what"": [""values""]
                    }
                ]
            ");

            // assert
            _workbook.GetSheetAt(0).GetRow(0).GetCell(1).StringCellValue.Should().Be("normal");
        }

        [Test]
        public void CanReferenceTheSourceSheetByName()
        {
            // act
            Run(@"
                [
                    {
                        ""kind"": ""copy"",
                        ""srcSheet"": ""TEST"",
                        ""src"": ""A1"",
                        ""dst"": ""B1"",
                        ""what"": [""values""]
                    }
                ]
            ");

            // assert
            _workbook.GetSheetAt(0).GetRow(0).GetCell(1).StringCellValue.Should().Be("normal");
        }

        [Test]
        public void CanReferenceTheDestinationSheetByName()
        {
            // act
            Run(@"
                [
                    {
                        ""kind"": ""copy"",
                        ""src"": ""A1"",
                        ""dstSheet"": ""TEST"",
                        ""dst"": ""B1"",
                        ""what"": [""values""]
                    }
                ]
            ");

            // assert
            _workbook.GetSheetAt(0).GetRow(0).GetCell(1).StringCellValue.Should().Be("normal");
        }

        [Test]
        public void CanCopyAnywhere()
        {
            // act
            Run(@"
                [
                    {
                        ""kind"": ""copy"",
                        ""src"": ""A1"",
                        ""dst"": ""F42""
                    }
                ]
            ");

            // assert
            _workbook.GetSheetAt(0).GetRow(41).GetCell(5).StringCellValue.Should().Be("normal");
        }

        [Test]
        public void CanCopyFormulas()
        {
            // act
            Run(@"
                [
                    {
                        ""kind"": ""copy"",
                        ""src"": ""A4"",
                        ""dst"": ""B1"",
                        ""what"": [""formulas""]
                    }
                ]
            ");

            // assert
            _workbook.GetSheetAt(0).GetRow(0).GetCell(1).CellFormula.Should().Be("RAND()");
        }

        [Test]
        public void ShiftsFormulas()
        {
            // act
            Run(@"
                [
                    {
                        ""kind"": ""copy"",
                        ""src"": ""A5"",
                        ""dst"": ""B42"",
                        ""what"": [""formulas""]
                    }
                ]
            ");

            // assert
            _workbook.GetSheetAt(0).GetRow(41).GetCell(1).CellFormula.Should().Be("B38");
        }

        [Test]
        public void CanCopyValues()
        {
            // act
            Run(@"
                [
                    {
                        ""kind"": ""copy"",
                        ""src"": ""A1"",
                        ""dst"": ""B1"",
                        ""what"": [""values""]
                    }
                ]
            ");

            // assert
            _workbook.GetSheetAt(0).GetRow(0).GetCell(1).StringCellValue.Should().Be("normal");
        }

        [Test]
        public void LeavesStylesIntactWhenCopyingValues()
        {
            // act
            Run(@"
                [
                    {
                        ""kind"": ""copy"",
                        ""src"": ""A1"",
                        ""dst"": ""A2:A3"",
                        ""what"": [""values""]
                    }
                ]
            ");

            // assert
            _workbook.GetSheetAt(0).GetRow(1).GetCell(0).CellStyle.GetFont(_workbook).IsBold.Should().BeTrue();
            _workbook.GetSheetAt(0).GetRow(2).GetCell(0).CellStyle.GetDataFormatString().Should().Be("d-mmm-yy");
        }

        [Test]
        public void CanCopyStyles()
        {
            // act
            Run(@"
                [
                    {
                        ""kind"": ""copy"",
                        ""src"": ""A2:A3"",
                        ""dst"": ""B1:B2"",
                        ""what"": [""styles""]
                    }
                ]
            ");

            // assert
            _workbook.GetSheetAt(0).GetRow(0).GetCell(1).CellStyle.GetFont(_workbook).IsBold.Should().BeTrue();
            _workbook.GetSheetAt(0).GetRow(1).GetCell(1).CellStyle.GetDataFormatString().Should().Be("d-mmm-yy");
        }

        [Test]
        public void LeavesFormulasAndValuesIntactWhenCopyingStyles()
        {
            // act
            Run(@"
                [
                    {
                        ""kind"": ""copy"",
                        ""src"": ""A1"",
                        ""dst"": ""A2:A4"",
                        ""what"": [""styles""]
                    }
                ]
            ");

            // assert
            _workbook.GetSheetAt(0).GetRow(1).GetCell(0).StringCellValue.Should().Be("bold");
            _workbook.GetSheetAt(0).GetRow(3).GetCell(0).CellFormula.Should().Be("RAND()");
        }

        [Test]
        public void CanCopyOneCellToARangeOfCells()
        {
            // act
            Run(@"
                [
                    {
                        ""kind"": ""copy"",
                        ""src"": ""A2"",
                        ""dst"": ""B1:B3""
                    }
                ]
            ");

            // assert
            _workbook.GetSheetAt(0).GetRow(0).GetCell(1).StringCellValue.Should().Be("bold");
            _workbook.GetSheetAt(0).GetRow(0).GetCell(1).CellStyle.GetFont(_workbook).IsBold.Should().BeTrue();
            _workbook.GetSheetAt(0).GetRow(1).GetCell(1).StringCellValue.Should().Be("bold");
            _workbook.GetSheetAt(0).GetRow(1).GetCell(1).CellStyle.GetFont(_workbook).IsBold.Should().BeTrue();
            _workbook.GetSheetAt(0).GetRow(2).GetCell(1).StringCellValue.Should().Be("bold");
            _workbook.GetSheetAt(0).GetRow(2).GetCell(1).CellStyle.GetFont(_workbook).IsBold.Should().BeTrue();
        }

        [Test]
        public void CanCopyARangeOfCellsToAnEquallySizedRangeOfCells()
        {
            // act
            Run(@"
                [
                    {
                        ""kind"": ""copy"",
                        ""src"": ""A1:A3"",
                        ""dst"": ""B1:B3""
                    }
                ]
            ");

            // assert
            _workbook.GetSheetAt(0).GetRow(0).GetCell(1).StringCellValue.Should().Be("normal");
            _workbook.GetSheetAt(0).GetRow(1).GetCell(1).StringCellValue.Should().Be("bold");
            _workbook.GetSheetAt(0).GetRow(1).GetCell(1).CellStyle.GetFont(_workbook).IsBold.Should().BeTrue();
            _workbook.GetSheetAt(0).GetRow(2).GetCell(1).NumericCellValue.Should().Be(43101);
            _workbook.GetSheetAt(0).GetRow(2).GetCell(1).CellStyle.GetDataFormatString().Should().Be("d-mmm-yy");
        }
    }
}
