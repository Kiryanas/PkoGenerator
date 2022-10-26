using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using NUnit.Framework;

namespace PkoGenerator.Tests
{
    [TestFixture]
    public class AccountingOperationTests
    {
        /// <summary>
        /// Парсинг бухгалтерской операции из входного файла.
        /// </summary>
        [Test]
        public void Parse()
        {
            var path = "doc\\TestCounterpartiesList.xlsx";
            AccountingOperation accountingOperation = null;
            using (var fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                using (var doc = SpreadsheetDocument.Open(fs, false))
                {
                    var workbookPart = doc.WorkbookPart;
                    var sstpart = workbookPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();
                    var sst = sstpart.SharedStringTable;
                    var worksheetPart = workbookPart.WorksheetParts.FirstOrDefault();
                    var sheet = worksheetPart.Worksheet;
                    var rows = sheet.Descendants<Row>();
                    try
                    { 
                        accountingOperation = AccountingOperation.Parse(sst, rows.First()); 
                    }
                    catch
                    { }
                }
            }
            Assert.IsNotNull(accountingOperation);
            Assert.AreEqual("Ромашка, ООО", accountingOperation.CounterpartyName);
            Assert.AreEqual(1000, accountingOperation.Amount);
        }
    }
}
