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
        public const string PkoTemplatePath = "SimpleOrderTemplate.xlsx";

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
            this.WriteToCell(pkoFilePath, "A", 2, accountingOperation.CounterpartyName);
            this.WriteToCell(pkoFilePath, "B", 2, accountingOperation.Amount.ToString());
            return true;
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
