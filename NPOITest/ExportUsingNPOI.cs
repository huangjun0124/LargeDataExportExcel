using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using Util;

namespace NPOITest
{
    class ExportUsingNPOI
    {
        public void DoTest()
        {
            var gdt = new GenerateDataTable();
            var table = gdt.GetNewTable();

            Stopwatch sw = new Stopwatch();
            sw.Start();
            var filename = System.IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory), DateTime.Now.ToString("yyyy-MM-dd_HH-mm-ss") + ".xlsx");
            ExportDataTableToExcel(table, filename);
            sw.Stop();
            Console.WriteLine($"Save used {sw.ElapsedMilliseconds} ms");
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
        }

        public static Tuple<bool, string> ExportDataTableToExcel(DataTable dt, string saveTopath)
        {
            bool result = false;
            string message = "";
            IWorkbook workbook = null;
            FileStream fs = null;
            IRow row = null;
            ISheet sheet = null;
            ICell cell = null;
            try
            {
                if (dt != null && dt.Rows.Count > 0)
                {
                    if (saveTopath.IndexOf(".xlsx", StringComparison.Ordinal) > 0) // 2007版本
                        workbook = new XSSFWorkbook();
                    else //if (saveTopath.IndexOf(".xls", StringComparison.Ordinal) > 0) // 2003版本
                        workbook = new HSSFWorkbook();
                    sheet = workbook.CreateSheet(dt.TableName);
                    int rowCount = dt.Rows.Count;//行数  
                    int columnCount = dt.Columns.Count;//列数  

                    //设置列头  
                    row = sheet.CreateRow(0);//excel第一行设为列头  
                    for (int c = 0; c < columnCount; c++)
                    {
                        cell = row.CreateCell(c);
                        cell.SetCellValue(dt.Columns[c].ColumnName);
                    }

                    //设置每行每列的单元格,  
                    for (int i = 0; i < rowCount; i++)
                    {
                        row = sheet.CreateRow(i + 1);
                        for (int j = 0; j < columnCount; j++)
                        {
                            cell = row.CreateCell(j);//excel第二行开始写入数据  
                            cell.SetCellValue(dt.Rows[i][j].ToString());
                        }
                    }
                    using (fs = File.OpenWrite(saveTopath))
                    {
                        workbook.Write(fs);//向打开的这个xls文件中写入数据  
                        result = true;
                    }
                }
                else
                {
                    message = "没有解析到数据！";
                }
                return new Tuple<bool, string>(result, message);
            }
            catch (Exception ex)
            {
                if (fs != null)
                {
                    fs.Close();
                }
                return new Tuple<bool, string>(false, ex.Message);
            }
        }
    }
}
