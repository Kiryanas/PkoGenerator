using DocumentFormat.OpenXml.Spreadsheet;

namespace PkoGenerator
{
    public class AccountingOperation
    {
        public string CounterpartyName;
        public decimal Amount;
        public static AccountingOperation Parse(Row row)
        {
            return new AccountingOperation();
        }
    }
}
