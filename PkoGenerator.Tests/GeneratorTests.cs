using NUnit.Framework;

namespace PkoGenerator.Tests
{
    [TestFixture]
    public class GeneratorTests
    {
        [Test]
        public void CreateNew()
        {
            new Generator(null);
        }

        [Test]
        public void GenerateSinglePko()
        {
            var generator = new Generator(Environment.CurrentDirectory);

            // Если введены корректные данные Контрагент и Сумма, формируется приходный кассовый ордер
            var operation = new AccountingOperation() { CounterpartyName = "ООО \"Ромашка\"", Amount = 35000 }; 
            Assert.IsTrue(generator.GenerateSinglePko(operation));

            // Если Контрагент имеет слишком длинное наименование, приходный кассовый ордер не формируется
            operation = new AccountingOperation() { CounterpartyName = "Очень длинное название ООО \"Ромашка\"", Amount = 35000 };
            Assert.IsFalse(generator.GenerateSinglePko(operation));

            // Если Сумма имеет отрицательное значение, приходный кассовый ордер не формируется
            operation = new AccountingOperation() { CounterpartyName = "ООО \"Ромашка\"", Amount = -2 };
            Assert.IsFalse(generator.GenerateSinglePko(operation));

            // Если Контрагент имеет пустое значение, приходный кассовый ордер не формируется
            operation = new AccountingOperation() { CounterpartyName = null, Amount = 1000 };
            Assert.IsFalse(generator.GenerateSinglePko(operation));

            // Если Сумма имеет значение, равное нулю, приходный кассовый ордер не формируется
            operation = new AccountingOperation() { CounterpartyName = "ООО \"Ромашка\"", Amount = 0 };
            Assert.IsFalse(generator.GenerateSinglePko(operation));
        }
    }
}