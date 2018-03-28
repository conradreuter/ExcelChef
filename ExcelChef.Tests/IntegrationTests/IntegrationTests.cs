using FluentAssertions;
using NPOI.SS.UserModel;
using NUnit.Framework;
using System.IO;

namespace ExcelChef.IntegrationTests
{
    [TestFixture]
    public class IntegrationTests
    {
        private static readonly string TemplateFile = Path.Combine(TestContext.CurrentContext.TestDirectory, @"IntegrationTests\template.xlsx");
        private static readonly string TemplateFileXls = Path.Combine(TestContext.CurrentContext.TestDirectory, @"IntegrationTests\template.xls");
        
        private IWorkbook _workbook;

        [TearDown]
        public void TearDown()
        {
            _workbook.Close();
        }

        [Test]
        public void CanHandleXlsx()
        {
            // act + assert
            RunProgram("[]");
        }

        [Test]
        public void CanHandleXls()
        {
            // act
            RunProgram("[]", xls: true);
        }

        [Test]
        public void DefaultsToUsingTheFirstWorksheet()
        {
            // act
            RunProgram(@"
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
            RunProgram(@"
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
            RunProgram(@"
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
            RunProgram(@"
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
            RunProgram(@"
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
            RunProgram(@"
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
            RunProgram(@"
                [
                    {
                        ""kind"": ""write"",
                        ""cell"": ""A3"",
                        ""value"": 43173
                    }
                ]
            ");

            // assert
            _workbook.GetSheetAt(0).GetRow(2).GetCell(0).CellStyle.GetDataFormatString().Should().Be("d-mmm-yy");
        }

        private void RunProgram(string instructions, bool xls = false)
        {
            // prepare input
            Stream input = new MemoryStream();
            TextWriter writer = new StreamWriter(input);
            writer.Write(instructions);
            writer.Flush();
            input.Seek(0, SeekOrigin.Begin);

            // run program
            Stream output = new MemoryStream();
            new Program
            {
                Input = input,
                Output = output,
                Template = new FileStream(xls ? TemplateFileXls : TemplateFile, FileMode.Open, FileAccess.Read),
            }.Run();

            // prepare output for assertions
            output.Seek(0, SeekOrigin.Begin);
            _workbook = WorkbookFactory.Create(output);
        }
    }
}
