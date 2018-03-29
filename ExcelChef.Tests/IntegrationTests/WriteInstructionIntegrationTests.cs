using FluentAssertions;
using NUnit.Framework;

namespace ExcelChef.IntegrationTests
{
    [TestFixture]
    public class WriteInstructionIntegrationTests : IntegrationTestsBase
    {
        [Test]
        public void DefaultsToUsingTheFirstWorksheet()
        {
            // act
            Run(@"
                [
                    {
                        ""kind"": ""write"",
                        ""cell"": ""A1"",
                        ""value"": ""value""
                    }
                ]
            ");

            // assert
            _workbook.GetSheetAt(0).GetRow(0).GetCell(0).StringCellValue.Should().Be("value");
        }

        [Test]
        public void CanReferenceWorksheetsByPosition()
        {
            // act
            Run(@"
                [
                    {
                        ""kind"": ""write"",
                        ""worksheet"": 1,
                        ""cell"": ""A1"",
                        ""value"": ""value""
                    }
                ]
            ");

            // assert
            _workbook.GetSheetAt(0).GetRow(0).GetCell(0).StringCellValue.Should().Be("value");
        }

        [Test]
        public void CanReferenceWorksheetsByName()
        {
            // act
            Run(@"
                [
                    {
                        ""kind"": ""write"",
                        ""worksheet"": ""TEST"",
                        ""cell"": ""A1"",
                        ""value"": ""value""
                    }
                ]
            ");

            // assert
            _workbook.GetSheetAt(0).GetRow(0).GetCell(0).StringCellValue.Should().Be("value");
        }

        [Test]
        public void LeavesStylingIntact()
        {
            // act
            Run(@"
                [
                    {
                        ""kind"": ""write"",
                        ""cell"": ""A2"",
                        ""value"": ""value""
                    }
                ]
            ");

            // assert
            _workbook.GetSheetAt(0).GetRow(1).GetCell(0).CellStyle.GetFont(_workbook).IsBold.Should().BeTrue();
        }

        [Test]
        public void CanWriteAnywhere()
        {
            // act
            Run(@"
                [
                    {
                        ""kind"": ""write"",
                        ""cell"": ""F42"",
                        ""value"": ""value""
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
                        ""cell"": ""A1"",
                        ""value"": 1337
                    }
                ]
            ");

            // assert
            _workbook.GetSheetAt(0).GetRow(0).GetCell(0).NumericCellValue.Should().Be(1337);
        }

        [Test]
        public void LeavesNumberFormatsIntact()
        {
            // act
            Run(@"
                [
                    {
                        ""kind"": ""write"",
                        ""cell"": ""A4"",
                        ""value"": 43173
                    }
                ]
            ");

            // assert
            _workbook.GetSheetAt(0).GetRow(2).GetCell(0).CellStyle.GetDataFormatString().Should().Be("d-mmm-yy");
        }
    }
}
