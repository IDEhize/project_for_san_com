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

            Word.Application wordApp = new Word.Application();
            Word.Document document = null;

            try
            {
                document = wordApp.Documents.Add();


                document.PageSetup.PaperSize = Word.WdPaperSize.wdPaperA4;
                document.PageSetup.Orientation = Word.WdOrientation.wdOrientPortrait;

                document.PageSetup.TopMargin = wordApp.CentimetersToPoints(2.0f);
                document.PageSetup.BottomMargin = wordApp.CentimetersToPoints(2.0f);
                document.PageSetup.LeftMargin = wordApp.CentimetersToPoints(3.0f);
                document.PageSetup.RightMargin = wordApp.CentimetersToPoints(1.5f);

                Word.Range entireDocument = document.Range();

                int columns = 3;
                Word.Table table = document.Tables.Add(document.Range(), rows, columns);
                table.Borders.Enable = 1;
                table.Cell(1, 1).Range.Text = "ком";
                table.Cell(1, 2).Range.Text = "дата";
                table.Cell(1, 3).Range.Text = "жильцы";

                float availableWidth = document.PageSetup.PageWidth - document.PageSetup.LeftMargin - document.PageSetup.RightMargin;
                float column1Width = availableWidth * 0.2f;
                float column2Width = availableWidth * 0.2f;
                float column3Width = availableWidth * 0.6f;

                table.Columns[1].Width = column1Width;
                table.Columns[2].Width = column2Width;
                table.Columns[3].Width = column3Width;
                for (int row = 2; row <= rows; row++)
                {
                    for (int col = 1; col <= columns; col++)
                    {
                        string cellText = $"Row {row}, Col {col}";
                        table.Cell(row, col).Range.Text = cellText;
                    }
                }
                foreach (Word.Table tbl in document.Tables)
                {
                    foreach (Word.Row row in tbl.Rows)
                    {
                        foreach (Word.Cell cell in row.Cells)
                        {
                            Word.Range cellRange = cell.Range;
                            cellRange.Font.Name = "Times New Roman";
                            cellRange.Font.Size = 14;
                            cellRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
                            cellRange.ParagraphFormat.LineSpacingRule = Word.WdLineSpacing.wdLineSpaceSingle;
                            cellRange.ParagraphFormat.SpaceBefore = 0;
                            cellRange.ParagraphFormat.SpaceAfter = 0;
                        }
                    }
                }
                document.SaveAs2(filePath);
                MessageBox.Show("Документ был создан и сохранен на рабочий стол");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка при создании документа: " + ex.Message);
            }
            finally
            {
                document?.Close();
                wordApp.Quit();
            }
        }
    }
}
