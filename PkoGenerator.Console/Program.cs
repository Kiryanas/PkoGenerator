namespace PkoGenerator.Console
{
    public class Program
    {
        public static void Main(string[] args)
        {
            if (args.Length != 2)
            {
                System.Console.WriteLine("Параметры:");
                System.Console.WriteLine("1. Путь до файла с бухгалтерскими операциями.");
                System.Console.WriteLine("2. Путь до папки с приходными кассовыми ордерами.");
                return;
            }

            var inputFilePath = args[0];
            if (!File.Exists(inputFilePath))
            {
                System.Console.WriteLine("Файл с бухгалтерскими операциями недоступен.");
                return;
            }

            var outputFolderPath = args[1];
            if (!Directory.Exists(outputFolderPath))
            {
                System.Console.WriteLine($"Создание папки {outputFolderPath}..."); 
                Directory.CreateDirectory(outputFolderPath);
                System.Console.WriteLine("Готово.");
            }

            var fileReader = new FileReader();

            var accontingOperations = new List<AccountingOperation>();
            try
            { 
                accontingOperations = fileReader.ReadXlsx<AccountingOperation>(inputFilePath); 
            }
            catch (Exception ex)
            {
                System.Console.WriteLine($"Не могу прочитать {inputFilePath}");
                System.Console.WriteLine(ex);
            }

            var generator = new Generator(outputFolderPath);
            foreach (var accontingOperation in accontingOperations)
            {
                try
                {
                    var pkoGenerated = generator.GenerateSinglePko(accontingOperation);
                    if (pkoGenerated)
                        System.Console.WriteLine($"ПКО успешно сгнерирован для {accontingOperation.CounterpartyName} ({accontingOperation.Amount}).");
                    else
                        System.Console.WriteLine($"ПКО не сгнерирован для {accontingOperation.CounterpartyName} ({accontingOperation.Amount}).");

                }
                catch (Exception ex)
                {
                    System.Console.WriteLine($"ПКО не сгнерирован для {accontingOperation.CounterpartyName} ({accontingOperation.Amount}).");
                    System.Console.WriteLine(ex);
                }
            }
        }
    }
}