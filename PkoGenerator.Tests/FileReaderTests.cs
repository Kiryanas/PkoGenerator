using NUnit.Framework;

namespace PkoGenerator.Tests
{
    [TestFixture]
    public class FileReaderTests
    {
        /// <summary>
        /// Основной кейс - чтение корректного файла с данными
        /// </summary>
        [Test]
        public void ReadXlsx()
        {
            var fileReader = new FileReader();
            var path = "doc\\TestCounterpartiesList.xlsx";
            var rows = fileReader.ReadXlsx<AccountingOperation>(path);
            Assert.NotNull(rows);
            Assert.AreEqual(3, rows.Count());
        }

        /// <summary>
        /// Пустой файл
        /// </summary>
        [Test]
        public void ReadEmptyXlsx()
        {
            var fileReader = new FileReader();
            var path = "doc\\EmptyFile.xlsx";
            var rows = fileReader.ReadXlsx<AccountingOperation>(path);
            Assert.IsEmpty(rows);
        }

        /// <summary>
        /// Несуществующий путь
        /// </summary>
        [Test]
        public void FileNotExist()
        {
            var fileReader = new FileReader();
            var path = "doc\\111.txt";
            Assert.Throws<FileNotFoundException>(() => { fileReader.ReadXlsx<AccountingOperation>(path); });
        }

        /// <summary>
        /// Неподходящий формат файла
        /// </summary>
        [Test]
        public void FileFormatNotSupported()
        {
            var fileReader = new FileReader();
            var path = "doc\\DifferentFormatFile.txt";
            Assert.Throws<ArgumentException>(() => { fileReader.ReadXlsx<AccountingOperation>(path); });
        }
    }
}
