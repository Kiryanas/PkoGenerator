using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.IO;
using System.Linq;
using System;
using DocumentFormat.OpenXml;

namespace PkoGenerator
{
    /// <summary>
    /// Генератор.
    /// </summary>
    public class Generator
    {
        /// <summary>
        /// Путь к файлу шаблона ПКО.
        /// </summary>
        public const string PkoTemplatePath = "OrderTemplate.xlsx";

        /// <summary>
        /// Папка для генерации ПКО.
        /// </summary>
        public string Destination { get; private set; }

        /// <summary>
        /// Формирование приходного кассового ордера.
        /// </summary>
        /// <param name="accountingOperation">Бухгалтерская операция.</param>
        /// <returns>true - ПКО успешно сформирован, false - иначе.</returns>

        public bool GenerateSinglePko(AccountingOperation accountingOperation)
        {
            if (accountingOperation == null)
                return false;
            if (accountingOperation.CounterpartyName == null)
                return false;
            if (accountingOperation.Amount <= 0)
                return false;
            if (accountingOperation.CounterpartyName.Length > 20)
                return false;

            var pkoName = string.Join(accountingOperation.CounterpartyName, accountingOperation.Amount.ToString(), ".xlsx");
            pkoName = pkoName.Replace(',', '_');
            pkoName = pkoName.Replace('"', '_');
            var pkoFilePath = Path.Combine(this.Destination, pkoName);
            File.Copy(PkoTemplatePath, pkoFilePath, true);
            this.WriteToCell(pkoFilePath, "K", 21, accountingOperation.CounterpartyName);
            this.WriteToCell(pkoFilePath, "A", 27, Math.Truncate(accountingOperation.Amount).ToString());
            var cents = Math.Truncate((accountingOperation.Amount - Math.Truncate(accountingOperation.Amount)) * 100);
            this.WriteToCell(pkoFilePath, "BC", 27, cents.ToString("00"));
            this.WriteToCell(pkoFilePath, "H", 25, this.DecimalToWords(accountingOperation.Amount));
            return true;
        }

        public string DecimalToWords(decimal number)
        {
            var integerPart = (int)Math.Truncate(number);
            var centsPart = (int)Math.Truncate((number - Math.Truncate(number)) * 100);
            var words = $"{integerPart} рублей {centsPart.ToString("00")} копеек";
            return words;
        }

        /// <summary>
        /// Записать в ячейку.
        /// </summary>
        /// <param name="path">Путь к файлу Excel.</param>
        /// <param name="columnName">Колонка.</param>
        /// <param name="rowIndex">Строка.</param>
        /// <param name="value">Значение.</param>
        public void WriteToCell(string path, string columnName, uint rowIndex, string value)
        {
            using (var fs = new FileStream(path, FileMode.Open, FileAccess.ReadWrite, FileShare.ReadWrite))
            {
                using (var doc = SpreadsheetDocument.Open(fs, true))
                {
                    var workbookPart = doc.WorkbookPart;
                    var sstpart = workbookPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();
                    if (sstpart == null)
                        return;
                    var sst = sstpart.SharedStringTable;
                    var worksheetPart = workbookPart.WorksheetParts.FirstOrDefault();
                    if (worksheetPart == null)
                        return;
                    var sheet = worksheetPart.Worksheet;
                    var cells = sheet.Descendants<Cell>();
                    var cell = cells.Where(x => x.CellReference.Value.Equals(columnName + rowIndex, StringComparison.InvariantCultureIgnoreCase)).FirstOrDefault();
                    if (cell == null)
                        return;
                    cell.CellValue = new CellValue(value);
                    cell.DataType = new EnumValue<CellValues>(CellValues.String);
                    sheet.Save();
                }
            }
        }

        /// <summary>
        /// Генератор.
        /// </summary>
        /// <param name="destination">Папка для генерации ПКО.</param>
        public Generator(string destination)
        {
            this.Destination = destination;
        }
    }
}
