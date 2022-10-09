using System.IO;

namespace PkoGenerator
{
    public class Generator
    {
        public const string TemplatePath = "SimpleOrderTemplate.xlsx";

        public string Destination { get; private set; }

        public bool GenerateSinglePko(AccountingOperation accountingOperation)
        {
            if (accountingOperation == null)
                return false;
            if (accountingOperation.CounterpartyName == null)
                return false;
            if (accountingOperation.Amount <= 0)
                return false;
            if (accountingOperation.CounterpartyName.Length>20)
                return false;

            var pkoName = string.Join(accountingOperation.CounterpartyName, accountingOperation.Amount.ToString(), ".xlsx");
            pkoName = pkoName.Replace(',', '_');
            pkoName = pkoName.Replace('"', '_');
            File.Copy(TemplatePath, Path.Combine(this.Destination, pkoName));

            return true;
        }

        public Generator(string destination)
        {
            this.Destination = destination;
        }
    }
}
