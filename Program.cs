using System;
using System.IO;
using System.Linq;
using Microsoft.Office.Interop.Word;

namespace XamlToDocxConverter
{
    class Program
    {
        static void Main(string[] args)
        {
            string projectPath = @"C:\проект";
            string outputDocxPath = @"C:\документ";

            try
            {
                ConvertXamlFilesToDocx(projectPath, outputDocxPath);
                Console.WriteLine("Конвертация завершена успешно!");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Ошибка: {ex.Message}");
            }
        }

        static void ConvertXamlFilesToDocx(string projectPath, string outputDocxPath)
        {
            Application wordApp = null;
            Document doc = null;

            try
            {
                wordApp = new Application();
                wordApp.Visible = false; // Скрыть Word во время работы
                doc = wordApp.Documents.Add();

                int listingNumber = 1;

                // Находим все XAML и XAML.cs файлы
                var xamlFiles = Directory.GetFiles(projectPath, "*.xaml", SearchOption.AllDirectories)
                    .Concat(Directory.GetFiles(projectPath, "*.xaml.cs", SearchOption.AllDirectories))
                    .OrderBy(f => f);

                foreach (string filePath in xamlFiles)
                {
                    Console.WriteLine($"Обрабатывается файл: {Path.GetFileName(filePath)}");
                    AddCodeListing(doc, filePath, listingNumber++);
                }

                // Создаем папку если не существует
                Directory.CreateDirectory(Path.GetDirectoryName(outputDocxPath));
                
                // Сохраняем документ
                doc.SaveAs2(outputDocxPath);
                Console.WriteLine($"Документ сохранен: {outputDocxPath}");
            }
            finally
            {
                // Важно: освобождаем ресурсы
                doc?.Close();
                wordApp?.Quit();
            }
        }

        static void AddCodeListing(Document doc, string filePath, int listingNumber)
        {
            string fileName = Path.GetFileName(filePath);
            string codeContent = File.ReadAllText(filePath);

            // Добавляем заголовок листинга - Times New Roman 14
            Paragraph heading = doc.Paragraphs.Add();
            heading.Range.Text = $"Листинг №{listingNumber} - Код \"{fileName}\"";
            heading.Range.Font.Name = "Times New Roman";
            heading.Range.Font.Size = 14;
            heading.Format.SpaceAfter = 0;
            heading.Range.ParagraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpace1pt5;
            heading.Range.InsertParagraphAfter();

            // Создаем таблицу для кода
            Table table = doc.Tables.Add(
                Range: doc.Paragraphs.Add().Range,
                NumRows: 1,
                NumColumns: 1
            );

            // Настраиваем таблицу
            table.Borders.OutsideLineStyle = WdLineStyle.wdLineStyleSingle;
            table.Borders.InsideLineStyle = WdLineStyle.wdLineStyleSingle;
            
            // Вставляем код в таблицу
            table.Cell(1, 1).Range.Text = codeContent;
            
            // Настраиваем шрифт для кода - Courier New 10
            table.Range.Font.Name = "Courier New";
            table.Range.Font.Size = 10;
            table.Range.ParagraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceSingle;
            
            // Добавляем отступ после таблицы
            doc.Paragraphs.Add();
        }

    }
}