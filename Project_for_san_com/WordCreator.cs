using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using Word = Microsoft.Office.Interop.Word;

namespace Project_for_san_com
{
    public class WordDocumentCreator
    {
        public void CreateAndFillWord(string fileName,int rows)
        {
            string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            string filePath = System.IO.Path.Combine(desktopPath, fileName);
            // Создание приложения Word (некий запуск самого ворда)
            Word.Application wordApp = new Word.Application();
            Word.Document document = null;

            try
            {
                // Создание нового документа в Ворде
                document = wordApp.Documents.Add();
                // Настройка документа
                document.PageSetup.PaperSize = Word.WdPaperSize.wdPaperA4; // А4
                document.PageSetup.Orientation = Word.WdOrientation.wdOrientPortrait; // Книжная ориентация
                // Установка полей в документа
                document.PageSetup.TopMargin = wordApp.CentimetersToPoints(2.5f);
                document.PageSetup.BottomMargin = wordApp.CentimetersToPoints(2.5f);
                document.PageSetup.LeftMargin = wordApp.CentimetersToPoints(3.0f);
                document.PageSetup.RightMargin = wordApp.CentimetersToPoints(1.5f);
                // Здесь код самого заполнения  файла
                int columns = 3;
                Word.Table table = document.Tables.Add(document.Range(), rows, columns);
                table.Borders.Enable = 1;

                table.Cell(1, 1).Range.Text = "ком"; // 1 столбец
                table.Cell(1, 2).Range.Text = "дата"; // 2 столбец
                table.Cell(1, 3).Range.Text = "жильцы"; // 3 столбец

                for (int row = 2; row <= rows; row++)
                {
                    for (int col = 1; col <= columns; col++)
                    {
                        string cellText = $"Row {row}, Col {col}"; // Пример текста
                        table.Cell(row, col).Range.Text = cellText;
                        Word.Range cellRange = table.Cell(row, col).Range;

                        // Настройка шрифта и выравнивания для ячейки
                        cellRange.Font.Name = "Times New Roman"; // Шрифт
                        cellRange.Font.Size = 14; // Размер шрифта
                        cellRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
                    }
                }
                // Сохранение документа
                document.SaveAs2(filePath);
                MessageBox.Show("Документ был создан и сохранен на рабочий стол");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка при создании документа: " + ex.Message);
            }
            finally
            {
                // Закрываем документ и приложение
                document?.Close();
                wordApp.Quit();
            }
        }
    }
}
