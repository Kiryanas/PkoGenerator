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

            // ���� ������� ���������� ������ ���������� � �����, ����������� ��������� �������� �����
            var operation = new AccountingOperation() { CounterpartyName = "��� \"�������\"", Amount = 35000 }; 
            Assert.IsTrue(generator.GenerateSinglePko(operation));

            // ���� ���������� ����� ������� ������� ������������, ��������� �������� ����� �� �����������
            operation = new AccountingOperation() { CounterpartyName = "����� ������� �������� ��� \"�������\"", Amount = 35000 };
            Assert.IsFalse(generator.GenerateSinglePko(operation));

            // ���� ����� ����� ������������� ��������, ��������� �������� ����� �� �����������
            operation = new AccountingOperation() { CounterpartyName = "��� \"�������\"", Amount = -2 };
            Assert.IsFalse(generator.GenerateSinglePko(operation));

            // ���� ���������� ����� ������ ��������, ��������� �������� ����� �� �����������
            operation = new AccountingOperation() { CounterpartyName = null, Amount = 1000 };
            Assert.IsFalse(generator.GenerateSinglePko(operation));

            // ���� ����� ����� ��������, ������ ����, ��������� �������� ����� �� �����������
            operation = new AccountingOperation() { CounterpartyName = "��� \"�������\"", Amount = 0 };
            Assert.IsFalse(generator.GenerateSinglePko(operation));
        }
    }
}