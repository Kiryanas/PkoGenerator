﻿using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace PkoGenerator
{
    /// <summary>
    /// Чтение входного файла Excel с бухгалтерскими операциями (контаргенты и суммы)
    /// </summary>
    public class FileReader
    {
        public List<T> ReadXlsx<T>(string path) where T : AccountingOperation
        {
            if (!File.Exists(path))
                throw new FileNotFoundException();
            if (new FileInfo(path).Extension != ".xlsx")
                throw new ArgumentException("Это не Excel!");
            var result = new List<T>();
            using (var fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                using (var doc = SpreadsheetDocument.Open(fs, false))
                {
                    var workbookPart = doc.WorkbookPart;
                    var sstpart = workbookPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();
                    if (sstpart == null) 
                        return result;
                    var sst = sstpart.SharedStringTable;
                    var worksheetPart = workbookPart.WorksheetParts.FirstOrDefault();
                    if (worksheetPart == null)
                        return result;
                    var sheet = worksheetPart.Worksheet;
                    var rows = sheet.Descendants<Row>();
                    foreach (var row in rows)
                    {
                        var accountingOperation = AccountingOperation.Parse(sst, row);
                        result.Add((T)accountingOperation);
                    }
                }
            }
            return result;       
        }
    }
}
