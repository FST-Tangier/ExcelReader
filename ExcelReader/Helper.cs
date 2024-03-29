using System;
using System.Collections.Generic;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelReader
{
    public class Helper
    {
        public static List<Article> Read(string relativePath)
        {
            List<Article> result = new List<Article>();
            string basePath = AppDomain.CurrentDomain.BaseDirectory;
            string filePath = Path.Combine(basePath, relativePath);
            Excel.Application excelApp = new Excel.Application();

            Excel.Workbook workbook = excelApp.Workbooks.Open(filePath);
            Excel.Worksheet worksheet = workbook.Sheets[1]; // Sélectionner la première feuille du classeur
            Excel.Range range = worksheet.UsedRange;

            int rowCount = range.Rows.Count;
            int colCount = range.Columns.Count;

            for (int ligne = 2; ligne <= rowCount; ligne++)
            {
                Article art = new Article(
                    range.Cells[ligne, 1].Value2.ToString(),
                    range.Cells[ligne, 2].Value2.ToString(),
                    range.Cells[ligne, 3].Value2.ToString());
                
                result.Add(art);
            }

            workbook.Close();
            excelApp.Quit();

            return result;
        }
    }
}



