using NUnit.Framework;

namespace PkoGenerator.Tests
{
    [TestFixture]
    public class FileReaderTests
    {
        [Test]
        public void ReadXlsx()
        {
            var fileReader = new FileReader();

            //Основной кейс - чтение корректного файла с данными
            var path = "doc\\TestCounterpartiesList.xlsx";
            var rows = fileReader.ReadXlsx<AccountingOperation>(path);
            Assert.NotNull(rows);
            Assert.AreEqual(3, rows.Count());

            //Пустой файл
            path = "doc\\EmptyFile.xlsx";
            rows = fileReader.ReadXlsx<AccountingOperation>(path);
            Assert.IsEmpty(rows);

            //Несуществующий путь
            path = "doc\\111.txt";
            Assert.Throws<FileNotFoundException>(() => { fileReader.ReadXlsx<AccountingOperation>(path); });

            //Неподходящий формат файла
            path = "doc\\DifferentFormatFile.txt";
            Assert.Throws<ArgumentException>(() => { fileReader.ReadXlsx<AccountingOperation>(path); });
        }
    }
}
