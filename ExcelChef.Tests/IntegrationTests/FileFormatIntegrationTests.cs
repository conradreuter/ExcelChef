using FluentAssertions;
using NPOI.SS.UserModel;
using NUnit.Framework;
using System.IO;

namespace ExcelChef.IntegrationTests
{
    [TestFixture]
    public class FileFormatIntegrationTests : IntegrationTestsBase
    {
        [Test]
        public void CanHandleXlsx()
        {
            // act + assert
            Run("[]");
        }

        [Test]
        public void CanHandleXls()
        {
            // act
            Run("[]", xls: true);
        }
    }
}
