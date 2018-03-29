using NPOI.SS.UserModel;
using NUnit.Framework;
using System.IO;

namespace ExcelChef.IntegrationTests
{
    public class IntegrationTestsBase
    {
        private static readonly string TemplateFile = Path.Combine(TestContext.CurrentContext.TestDirectory, @"IntegrationTests\template.xlsx");
        private static readonly string TemplateFileXls = Path.Combine(TestContext.CurrentContext.TestDirectory, @"IntegrationTests\template.xls");

        protected IWorkbook _workbook;

        [TearDown]
        public void TearDown()
        {
            _workbook.Close();
        }

        protected void Run(string instructions, bool xls = false)
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
