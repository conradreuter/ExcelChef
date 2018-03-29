using FluentAssertions;
using NUnit.Framework;

namespace ExcelChef.IntegrationTests
{
    [TestFixture]
    public class WriteInstructionIntegrationTests : IntegrationTestsBase
    {
        [Test]
        public void DefaultsToUsingTheFirstSheet()
        {
            // act
            Run(@"
                [
                    {
                        ""kind"": ""write"",
                        ""dst"": ""A1"",
                        ""values"": [""value""]
                    }
                ]
            ");

            // assert
            _workbook.GetSheetAt(0).GetRow(0).GetCell(0).StringCellValue.Should().Be("value");
        }

        [Test]
        public void CanReferenceTheSheetByPosition()
        {
            // act
            Run(@"
                [
                    {
                        ""kind"": ""write"",
                        ""dstSheet"": 1,
                        ""dst"": ""A1"",
                        ""values"": [""value""]
                    }
                ]
            ");

            // assert
            _workbook.GetSheetAt(0).GetRow(0).GetCell(0).StringCellValue.Should().Be("value");
        }

        [Test]
        public void CanReferenceTheSheetByName()
        {
            // act
            Run(@"
                [
                    {
                        ""kind"": ""write"",
                        ""dstSheet"": ""TEST"",
                        ""dst"": ""A1"",
                        ""values"": [""value""]
                    }
                ]
            ");

            // assert
            _workbook.GetSheetAt(0).GetRow(0).GetCell(0).StringCellValue.Should().Be("value");
        }

        [Test]
        public void CanWriteAnywhere()
        {
            // act
            Run(@"
                [
                    {
                        ""kind"": ""write"",
                        ""dst"": ""F42"",
                        ""values"": [""value""]
                    }
                ]
            ");

            // assert
            _workbook.GetSheetAt(0).GetRow(41).GetCell(5).StringCellValue.Should().Be("value");
        }

        [Test]
        public void CanWriteNumbers()
        {
            // act
            Run(@"
                [
                    {
                        ""kind"": ""write"",
                        ""dst"": ""A1"",
                        ""values"": [1337]
                    }
                ]
            ");

            // assert
            _workbook.GetSheetAt(0).GetRow(0).GetCell(0).NumericCellValue.Should().Be(1337);
        }

        [Test]
        public void CanWriteMultipleValuesAtOnce()
        {
            // act
            Run(@"
                [
                    {
                        ""kind"": ""write"",
                        ""dst"": ""A1:B2"",
                        ""values"": [""value00"", ""value01"", ""value10"", ""value11""]
                    }
                ]
            ");

            // assert
            _workbook.GetSheetAt(0).GetRow(0).GetCell(0).StringCellValue.Should().Be("value00");
            _workbook.GetSheetAt(0).GetRow(0).GetCell(1).StringCellValue.Should().Be("value01");
            _workbook.GetSheetAt(0).GetRow(1).GetCell(0).StringCellValue.Should().Be("value10");
            _workbook.GetSheetAt(0).GetRow(1).GetCell(1).StringCellValue.Should().Be("value11");
        }

        [Test]
        public void LeavesStylesIntact()
        {
            // act
            Run(@"
                [
                    {
                        ""kind"": ""write"",
                        ""dst"": ""A2:A3"",
                        ""values"": [""value"", 43173]
                    }
                ]
            ");

            // assert
            _workbook.GetSheetAt(0).GetRow(1).GetCell(0).CellStyle.GetFont(_workbook).IsBold.Should().BeTrue();
            _workbook.GetSheetAt(0).GetRow(2).GetCell(0).CellStyle.GetDataFormatString().Should().Be("d-mmm-yy");
        }
    }
}
