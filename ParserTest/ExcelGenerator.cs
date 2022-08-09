using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;
using System.Reflection;

namespace ParserTest
{
    class ExcelGenerator
    {
        static List<string> Characteristics = new List<string>() { "Title", "Brand", "Id", "Feedbacks", "Price" };
        public static void DisplayInExcel(List<string> keyWords, List<Responce> responce, bool isVisible)
        {
            var book = OpenDocument(isVisible);
            CreateWorkSheets(book, keyWords);
            foreach (var resp in responce) 
            {
                GenerateCells(resp.Data, GetWorksheetByName(book, resp.Metadata.Name), book);
            }


        }

        public static Excel._Workbook OpenDocument(bool isVisible) 
        {
            var excelApp = new Excel.Application();
            excelApp.DisplayAlerts = false;
            excelApp.Visible = isVisible;
            var book = excelApp.Workbooks.Open(Path.GetFullPath("../../../ExcelAndKeys/Result"), ReadOnly: false);
            return book;
        }

        public static Excel._Workbook CreateWorkSheets(Excel._Workbook workbook, List<string> ketWords) 
        {
            try
            {
                Excel._Worksheet _workSheet = null;
                foreach (Excel.Worksheet worksheet in workbook.Worksheets)
                {
                    if (workbook.Worksheets.Count == 1)
                    {
                        break;
                    }
                    worksheet.Delete();
                }

                for (int i = 1; i < ketWords.Count + 1; i++)
                {
                    if (i == 1)
                    {
                        _workSheet = (Excel.Worksheet)workbook.ActiveSheet;
                        _workSheet.Name = ketWords[i - 1];
                        continue;
                    }
                    workbook.Worksheets.Add();
                    _workSheet = (Excel.Worksheet)workbook.ActiveSheet;
                    _workSheet.Name = ketWords[i - 1];

                }

                workbook.Save();
            }
            catch (Exception ex)
            {
                Console.WriteLine("Something wrong");
                Console.WriteLine(ex.Message);
            }

            return workbook;

        }

        public static void GenerateCells(ProductsList productsList, Excel._Worksheet workSheet, Excel._Workbook workbook) 
        {
            try 
            {
                for (int q = 1; q < Characteristics.Count + 1; q++)
                {
                    workSheet.Cells[1, q] = Characteristics[q - 1];
                }

                for (int c = 1; c < productsList.Products.Count + 1; c++)
                {

                    for (int i = 1; i < typeof(Product).GetProperties().Length + 1; i++)
                    {
                        var prop = typeof(Product).GetProperties()[i - 1].GetValue(productsList.Products[c - 1]).ToString();
                        workSheet.Cells[c + 1, i] = prop;
                    }
                }

                Excel.Range usedrange = workSheet.UsedRange;
                usedrange.Columns.AutoFit();
            }
            catch (Exception ex)
            {
                Console.WriteLine("Something wrong");
                Console.WriteLine(ex.Message);
            }

            workbook.Save();

        }

        public static Excel._Worksheet GetWorksheetByName(Excel._Workbook workbook, string metadataName)
        {

            var name = char.ToUpper(metadataName[0]) + metadataName.Substring(1);
            Excel.Worksheet worksheet = (Excel.Worksheet)workbook.Worksheets[$"{name}"];
            return worksheet;
        }
    }
}
