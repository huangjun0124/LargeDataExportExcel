using System;
using System.Diagnostics;
using ClosedXML.Excel;
using Util;

namespace ExportUsingClosedXML
{
    public class ClosedXmlTest
    {
        public void DoTest_GenerateDirectly()
        {
            var gdt = new GenerateDataTable();
            var table = gdt.GetNewTable();

            Stopwatch sw = new Stopwatch();
            sw.Start();
            using (var workbook = new XLWorkbook())
            {
                var filePath = System.IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory), DateTime.Now.ToString("yyyy-MM-dd_HH-mm-ss")+ ".xlsx");
                var worksheet = workbook.Worksheets.Add(table, table.TableName);
                sw.Stop();
                Console.WriteLine($"Add workbook used {sw.ElapsedMilliseconds} ms");
                sw.Restart();
                workbook.SaveAs(filePath);
            }
            sw.Stop();
            Console.WriteLine($"Save used {sw.ElapsedMilliseconds} ms");
        }

        public void DoTest_GenerateThenAttach()
        {
            var gdt = new GenerateDataTable();
            var table = gdt.GetNewTable(1);

            Stopwatch sw = new Stopwatch();
            sw.Start();
            var filePath = System.IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory), DateTime.Now.ToString("yyyy-MM-dd_HH-mm-ss") + ".xlsx");
            int looRowCount = 8000;
            using (var workbook = new XLWorkbook())
            {
                workbook.Worksheets.Add(gdt.GetTableWithNRows(table, looRowCount), table.TableName);
                workbook.SaveAs(filePath);
            }

            int i = 2;
            int loop = 300000 / looRowCount;
            while (i <= loop)
            {
                using (var workbook = new XLWorkbook(filePath))
                {
                    IXLWorksheet Worksheet = workbook.Worksheet(table.TableName);
                    int NumberOfLastRow = Worksheet.LastRowUsed().RowNumber();
                    IXLCell CellForNewData = Worksheet.Cell(NumberOfLastRow + 1, 1);
                    CellForNewData.InsertTable(gdt.GetTableWithNRows(table, looRowCount));
                    if (i == loop)
                    {
                        Worksheet.Columns().AdjustToContents();
                    }
                    workbook.Save();
                    Console.WriteLine($"Loop {i} work done...");
                    i++;
                }

                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();
            }

            sw.Stop();
            Console.WriteLine($"Save all rows used {sw.ElapsedMilliseconds} ms");
        }

    }
}
