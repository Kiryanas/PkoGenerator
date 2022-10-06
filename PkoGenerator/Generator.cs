using DocumentFormat.OpenXml.Packaging;
using System;

namespace PkoGenerator
{
    public class Generator
    {
        public byte[] GenerateSinglePko(string counterpartyName, decimal amount)
        {
            if (counterpartyName == null)
                return null;
            if (amount <= 0)
                return null;
            if (counterpartyName.Length>20)
                return null;
            var pko = new byte[counterpartyName.Length];
            return pko;
        }
    }
}
