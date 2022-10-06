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

            // ���� ������� ���������� ������ ���������� � �����, ����������� ��������� �������� �����
            var counterpartyName = "��� \"�������\"";
            decimal amount = 35000;
            var pko = generator.GenerateSinglePko(counterpartyName, amount);
            Assert.NotNull(pko);

            // ���� ���������� ����� ������� ������� ������������, ��������� �������� ����� �� �����������
            counterpartyName = "����� ������� �������� ��� \"�������\"";
            amount = 35000;
            pko = generator.GenerateSinglePko(counterpartyName, amount);
            Assert.Null(pko);

            // ���� ����� ����� ������������� ��������, ��������� �������� ����� �� �����������
            counterpartyName = "��� \"�������\"";
            amount = -2;
            pko = generator.GenerateSinglePko(counterpartyName, amount);
            Assert.Null(pko);

            // ���� ���������� ����� ������ ��������, ��������� �������� ����� �� �����������
            counterpartyName = null;
            amount = 1000;
            pko = generator.GenerateSinglePko(counterpartyName, amount);
            Assert.Null(pko);

            // ���� ����� ����� ��������, ������ ����, ��������� �������� ����� �� �����������
            counterpartyName = "��� \"�������\"";
            amount = 0;
            pko = generator.GenerateSinglePko(counterpartyName, amount);
            Assert.Null(pko);
        }
    }
}