using System;
using System.Collections.Generic;
using System.IO;
using ClosedXML.Excel;

class Program
{
    static void Main()
    {
        var filePath = "input/input.xlsx";
        var outputFilePath = "output/output.txt";

        using (var workbook = new XLWorkbook(filePath))
        {
            var worksheet = workbook.Worksheet(1); 

            int rowCount = worksheet.LastRowUsed().RowNumber();
            int colCount = worksheet.LastColumnUsed().ColumnNumber();

            Dictionary<int, Dictionary<string, int>> columnWordCounts = new Dictionary<int, Dictionary<string, int>>();


            for (int col = 1; col <= colCount; col++)
            {

                if (col == GetColumnNumber("A") || col == GetColumnNumber("AR") || col == GetColumnNumber("AS") || col == GetColumnNumber("AQ"))
                {
                    continue;
                }

                columnWordCounts[col] = new Dictionary<string, int>();

                for (int row = 2; row <= rowCount; row++)
                {
                    var cellValue = worksheet.Cell(row, col).GetString();
                    if (!string.IsNullOrEmpty(cellValue))
                    {
                        if (!columnWordCounts[col].ContainsKey(cellValue))
                        {
                            columnWordCounts[col][cellValue] = 0;
                        }
                        columnWordCounts[col][cellValue]++;
                    }
                }
            }

            Dictionary<string, string> aqContents = new Dictionary<string, string>();
            for (int row = 2; row <= rowCount; row++) 
            {
               var aqValue = worksheet.Cell(row, "AQ").GetString();
                if (!string.IsNullOrEmpty(aqValue))
                {
                    var eValue = worksheet.Cell(row, "E").GetString();
                    aqContents[row.ToString()] = $"{eValue} / {aqValue}";
                }
            }


            using (StreamWriter writer = new StreamWriter(outputFilePath))
            {
                foreach (var colEntry in columnWordCounts)
                {
                    writer.WriteLine($"Column {GetColumnName(colEntry.Key)}:");
                    foreach (var wordEntry in colEntry.Value)
                    {
                        writer.WriteLine($"  '{wordEntry.Key}': {wordEntry.Value} 人");
                    }
                }

                writer.WriteLine("Special Output (E / AQ):");
                foreach (var kvp in aqContents)
                {
                    if (!string.IsNullOrEmpty(kvp.Value))
                    {
                        writer.WriteLine($"{kvp.Key} / {kvp.Value}");
                    }
                }
            }

            Console.WriteLine($"Results have been written to {outputFilePath}");
        }
    }

    // 列名を列番号に変換するメソッド
    static int GetColumnNumber(string columnName)
    {
        int number = 0;
        for (int i = 0; i < columnName.Length; i++)
        {
            number = number * 26 + (columnName[i] - 'A' + 1);
        }
        return number;
    }

    // 列番号を列名に変換するメソッド
    static string GetColumnName(int columnNumber)
    {
        string columnName = "";
        while (columnNumber > 0)
        {
            int modulo = (columnNumber - 1) % 26;
            columnName = Convert.ToChar('A' + modulo) + columnName;
            columnNumber = (columnNumber - modulo) / 26;
        }
        return columnName;
    }
}
