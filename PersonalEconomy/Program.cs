using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;

namespace PersonalEconomy
{
    class Program
    {
        static void Main(string[] args)
        {
            Application xlApp = new Application();
            Workbook xlWorkBook = null;
            Worksheet xlWorkSheet;
            try
            {
                xlWorkBook = xlApp.Workbooks.Open(@"C:\Users\nocarbri\OneDrive\Documentos\Economy\2018\Transaksjoner.xlsx");
                xlWorkSheet = (Worksheet)xlWorkBook.Worksheets.get_Item(1);

                List<Mapping> costMapping = new List<Mapping>();
                costMapping = GetMapping(xlWorkBook, 3);
                List<Mapping> incomeMapping = new List<Mapping>();
                incomeMapping = GetMapping(xlWorkBook, 2);

                Range rng = xlWorkSheet.UsedRange;
                foreach (Range rangeRow in xlWorkSheet.UsedRange.Rows)
                {
                    if (rangeRow.Row != 1 && rng.Cells[rangeRow.Row, 1] != null)
                    {
                        string description = rng.Cells[rangeRow.Row, 4].Value2;
                        var costCategory = costMapping.Where(map => description.ToLower().Contains(map.Keyword.ToLower()));
                        var incomeCategory = incomeMapping.Where(map => description.ToLower().Contains(map.Keyword.ToLower()));
                        if (costCategory != null && costCategory.Any())
                        {
                            string category = costCategory.FirstOrDefault().Value;
                            rng.Cells[rangeRow.Row, 10] = category;
                            Console.WriteLine("Description: " + rng.Cells[rangeRow.Row, 4].Value2 + ". Setting value: " + category);
                        }
                        else if (incomeCategory != null && incomeCategory.Any())
                        {
                            string category = incomeCategory.FirstOrDefault().Value;
                            rng.Cells[rangeRow.Row, 10] = category;
                            Console.WriteLine("Description: " + rng.Cells[rangeRow.Row, 4].Value2 + ". Setting value: " + category);
                        }
                        else
                        {
                            Console.WriteLine("Description: " + rng.Cells[rangeRow.Row, 4].Value2);
                        }
                    }
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message + e.StackTrace);
            }
            xlApp.Quit();
            Console.ReadLine();
        }

        private static List<Mapping> GetMapping(Workbook xlWorkBook, int sheetIndex)
        {
            List<Mapping> list = new List<Mapping>();
            try
            {
                Worksheet xlWorkSheet = (Worksheet)xlWorkBook.Worksheets.get_Item(sheetIndex);
                Range rng = xlWorkSheet.UsedRange;
                foreach (Range rangeRow in xlWorkSheet.UsedRange.Rows)
                {
                    if (rangeRow.Row != 1 && rng.Cells[rangeRow.Row, 1] != null)
                    {
                        Console.WriteLine("Mapping: " + rng.Cells[rangeRow.Row, 1].Value2 + " - " + rng.Cells[rangeRow.Row, 2].Value2);
                        list.Add(new Mapping()
                        {
                            Keyword = rng.Cells[rangeRow.Row, 1].Value2,
                            Value = rng.Cells[rangeRow.Row, 2].Value2
                        });
                    }
                }
            }
            catch (Exception)
            {

            }
            return list;
        }
    }

    internal class Mapping
    {
        public string Keyword { get; set; }
        public string Value { get; set; }
    }
}
