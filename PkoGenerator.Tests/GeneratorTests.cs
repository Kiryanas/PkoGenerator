using NUnit.Framework;

namespace PkoGenerator.Tests
{
    [TestFixture]
    public class GeneratorTests
    {
        [Test]
        public void CreateNew()
        {
            new Generator();
        }

        [Test]
        public void GenerateSinglePko()
        {
            var generator = new Generator();

            // Если введены корректные данные Контрагент и Сумма, формируется приходный кассовый ордер
            var counterpartyName = "ООО \"Ромашка\"";
            decimal amount = 35000;
            var pko = generator.GenerateSinglePko(counterpartyName, amount);
            Assert.NotNull(pko);

            // Если Контрагент имеет слишком длинное наименование, приходный кассовый ордер не формируется
            counterpartyName = "Очень длинное название ООО \"Ромашка\"";
            amount = 35000;
            pko = generator.GenerateSinglePko(counterpartyName, amount);
            Assert.Null(pko);

            // Если Сумма имеет отрицательное значение, приходный кассовый ордер не формируется
            counterpartyName = "ООО \"Ромашка\"";
            amount = -2;
            pko = generator.GenerateSinglePko(counterpartyName, amount);
            Assert.Null(pko);

            // Если Контрагент имеет пустое значение, приходный кассовый ордер не формируется
            counterpartyName = null;
            amount = 1000;
            pko = generator.GenerateSinglePko(counterpartyName, amount);
            Assert.Null(pko);

            // Если Сумма имеет значение, равное нулю, приходный кассовый ордер не формируется
            counterpartyName = "ООО \"Ромашка\"";
            amount = 0;
            pko = generator.GenerateSinglePko(counterpartyName, amount);
            Assert.Null(pko);
        }
    }
}