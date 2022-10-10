using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using NUnit.Framework;

namespace PkoGenerator.Tests
{
    [TestFixture]
    public class GeneratorTests
    {
        /// <summary>
        /// Создание пустого ПКО
        /// </summary>
        [Test]
        public void CreateNew()
        {
            new Generator(null);
        }

        /// <summary>
        /// Создание ПКО для одного контрагента
        /// </summary>
        [Test]
        public void GenerateSinglePko()
        {
            var generator = new Generator(Environment.CurrentDirectory);
            var operation = new AccountingOperation() { CounterpartyName = "ООО \"Ромашка\"", Amount = 35000 };
            Assert.IsTrue(generator.GenerateSinglePko(operation));
            FileAssert.Exists("35000ООО _Ромашка_.xlsx");
            var name = string.Empty;
            var amount = string.Empty;
            using (var fs = new FileStream("35000ООО _Ромашка_.xlsx", FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                using (var doc = SpreadsheetDocument.Open(fs, false))
                {
                    var workbookPart = doc.WorkbookPart;
                    var sstpart = workbookPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();
                    var sst = sstpart.SharedStringTable;
                    var worksheetPart = workbookPart.WorksheetParts.FirstOrDefault();
                    var sheet = worksheetPart.Worksheet;
                    var cells = sheet.Descendants<Cell>();
                    var counterpartyNameCell = cells.Where(x => x.CellReference.Value == "A2").First();
                    var amountCell = cells.Where(x => x.CellReference.Value == "B2").First();
                    name = counterpartyNameCell.CellValue.InnerText;
                    amount = amountCell.CellValue.InnerText;
                }
            }
            Assert.AreEqual("ООО \"Ромашка\"", name);
            Assert.AreEqual("35000", amount);
        }

        /// <summary>
        ///  Если Контрагент имеет слишком длинное наименование, ПКО не формируется
        /// </summary>
        [Test]
        public void TooLongCounterpartyName()
        {
            var generator = new Generator(Environment.CurrentDirectory);
            var operation = new AccountingOperation() { CounterpartyName = "Очень длинное название ООО \"Ромашка\"", Amount = 35000 };
            Assert.IsFalse(generator.GenerateSinglePko(operation));

        }

        /// <summary>
        ///  Если Сумма имеет отрицательное значение, ПКО не формируется
        /// </summary>
        [Test]
        public void NegativeAmount()
        {
            var generator = new Generator(Environment.CurrentDirectory);
            var operation = new AccountingOperation() { CounterpartyName = "ООО \"Ромашка\"", Amount = -2 };
            Assert.IsFalse(generator.GenerateSinglePko(operation));
        }

        /// <summary>
        ///  Если Контрагент имеет пустое значение, ПКО не формируется
        /// </summary>
        [Test]
        public void NullCounterpartyName()
        {
            var generator = new Generator(Environment.CurrentDirectory);
            var operation = new AccountingOperation() { CounterpartyName = null, Amount = 1000 };
            Assert.IsFalse(generator.GenerateSinglePko(operation));
        }

        /// <summary>
        ///  Если Сумма имеет значение, равное нулю, ПКО не формируется
        /// </summary>
        [Test]
        public void ZeroAmount()
        {
            var generator = new Generator(Environment.CurrentDirectory);
            var operation = new AccountingOperation() { CounterpartyName = "ООО \"Ромашка\"", Amount = 0 };
            Assert.IsFalse(generator.GenerateSinglePko(operation));
        }

        /// <summary>
        ///  Запись в ячейку файла Excel
        /// </summary>
        [Test]
        public void WriteToCell()
        {
            var generator = new Generator(Environment.CurrentDirectory);
            var pkoFilePath = "TestExcelFile.xlsx";
            File.Copy(Generator.TemplatePath, pkoFilePath, true);
            Assert.DoesNotThrow(() => { generator.WriteToCell(pkoFilePath, "A", 2, "ООО \"Ромашка\""); });
            var str = string.Empty;
            using (var fs = new FileStream(pkoFilePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                using (var doc = SpreadsheetDocument.Open(fs, false))
                {
                    var workbookPart = doc.WorkbookPart;
                    var sstpart = workbookPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();
                    var sst = sstpart.SharedStringTable;
                    var worksheetPart = workbookPart.WorksheetParts.FirstOrDefault();
                    var sheet = worksheetPart.Worksheet;
                    var cells = sheet.Descendants<Cell>();
                    var cell = cells.Where(x => x.CellReference.Value == "A2").First();
                    str = cell.CellValue.InnerText;
                }
            }
            Assert.AreEqual("ООО \"Ромашка\"", str);
        }

    }
}