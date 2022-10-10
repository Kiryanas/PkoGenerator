using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using NUnit.Framework;

namespace PkoGenerator.Tests
{
    [TestFixture]
    public class GeneratorTests
    {
        /// <summary>
        /// �������� ������� ���
        /// </summary>
        [Test]
        public void CreateNew()
        {
            new Generator(null);
        }

        /// <summary>
        /// �������� ��� ��� ������ �����������
        /// </summary>
        [Test]
        public void GenerateSinglePko()
        {
            var generator = new Generator(Environment.CurrentDirectory);
            var operation = new AccountingOperation() { CounterpartyName = "��� \"�������\"", Amount = 35000 };
            Assert.IsTrue(generator.GenerateSinglePko(operation));
            FileAssert.Exists("35000��� _�������_.xlsx");
            var name = string.Empty;
            var amount = string.Empty;
            using (var fs = new FileStream("35000��� _�������_.xlsx", FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
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
            Assert.AreEqual("��� \"�������\"", name);
            Assert.AreEqual("35000", amount);
        }

        /// <summary>
        ///  ���� ���������� ����� ������� ������� ������������, ��� �� �����������
        /// </summary>
        [Test]
        public void TooLongCounterpartyName()
        {
            var generator = new Generator(Environment.CurrentDirectory);
            var operation = new AccountingOperation() { CounterpartyName = "����� ������� �������� ��� \"�������\"", Amount = 35000 };
            Assert.IsFalse(generator.GenerateSinglePko(operation));

        }

        /// <summary>
        ///  ���� ����� ����� ������������� ��������, ��� �� �����������
        /// </summary>
        [Test]
        public void NegativeAmount()
        {
            var generator = new Generator(Environment.CurrentDirectory);
            var operation = new AccountingOperation() { CounterpartyName = "��� \"�������\"", Amount = -2 };
            Assert.IsFalse(generator.GenerateSinglePko(operation));
        }

        /// <summary>
        ///  ���� ���������� ����� ������ ��������, ��� �� �����������
        /// </summary>
        [Test]
        public void NullCounterpartyName()
        {
            var generator = new Generator(Environment.CurrentDirectory);
            var operation = new AccountingOperation() { CounterpartyName = null, Amount = 1000 };
            Assert.IsFalse(generator.GenerateSinglePko(operation));
        }

        /// <summary>
        ///  ���� ����� ����� ��������, ������ ����, ��� �� �����������
        /// </summary>
        [Test]
        public void ZeroAmount()
        {
            var generator = new Generator(Environment.CurrentDirectory);
            var operation = new AccountingOperation() { CounterpartyName = "��� \"�������\"", Amount = 0 };
            Assert.IsFalse(generator.GenerateSinglePko(operation));
        }

        /// <summary>
        ///  ������ � ������ ����� Excel
        /// </summary>
        [Test]
        public void WriteToCell()
        {
            var generator = new Generator(Environment.CurrentDirectory);
            var pkoFilePath = "TestExcelFile.xlsx";
            File.Copy(Generator.TemplatePath, pkoFilePath, true);
            Assert.DoesNotThrow(() => { generator.WriteToCell(pkoFilePath, "A", 2, "��� \"�������\""); });
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
            Assert.AreEqual("��� \"�������\"", str);
        }

    }
}