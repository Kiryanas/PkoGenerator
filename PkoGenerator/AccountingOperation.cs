using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Globalization;
using System.Linq;

namespace PkoGenerator
{
    /// <summary>
    /// Бухгалтерская операция
    /// </summary>
    public class AccountingOperation
    {
        /// <summary>
        /// Название контрагента
        /// </summary>
        public string CounterpartyName;

        /// <summary>
        /// Сумма операции
        /// </summary>
        public decimal Amount;

        /// <summary>
        /// Конвертация строки файла Excel в бухгалтерскую операцию.
        /// </summary>
        /// <param name="sharedStringTable">Таблица строковых значений файла Excel.</param>
        /// <param name="row">Строка файла Excel.</param>
        /// <returns>Бухгалтерская операция.</returns>
        public static AccountingOperation Parse(SharedStringTable sharedStringTable, Row row)
        {
            var accountingOperation = new AccountingOperation();
            var cells = row.Elements<Cell>().ToList();
            var nameCell = cells[0];
            if ((nameCell.DataType != null) && (nameCell.DataType == CellValues.SharedString))
            {
                var ssid = int.Parse(nameCell.CellValue.Text);
                var str = sharedStringTable.ChildElements[ssid].InnerText;
                accountingOperation.CounterpartyName = str;
            }
            var amountCell = cells[1];
            if (amountCell.DataType == null)
            {
                var str = amountCell.CellValue.Text;
                accountingOperation.Amount = decimal.Parse(str, CultureInfo.InvariantCulture);
            }
            return accountingOperation;
        }
    }
}
