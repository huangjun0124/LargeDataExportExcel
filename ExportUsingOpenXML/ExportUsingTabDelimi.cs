using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Util;

namespace ExportUsingClosedXML
{
    /// <summary>
    /// comes from : https://www.cnblogs.com/brambling/p/6854731.html
    /// Not working, the file can not be opened
    /// </summary>
    class ExportUsingTabDelimi
    {
        public void DoTest()
        {
            var gdt = new GenerateDataTable();
            var dt = gdt.GetNewTable();

            Stopwatch stopwatch = new Stopwatch();
            stopwatch.Start();
            var filename = System.IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory), DateTime.Now.ToString("yyyy-MM-dd_HH-mm-ss") + ".xlsx");

            //创建文件
            FileStream file = new FileStream(filename, FileMode.CreateNew, FileAccess.Write);

            //以指定的字符编码向指定的流写入字符
            StreamWriter sw = new StreamWriter(file, Encoding.UTF8);

            //写入标题
            for (int i = 0; i < dt.Columns.Count; i++)
            {
                sw.Write(dt.Columns[i].ColumnName.ToString() + "\t");
            }
            //加入换行字符串
            sw.Write(Environment.NewLine);

            //写入内容
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                for (int j = 0; j < dt.Columns.Count; j++)
                {
                    sw.Write(dt.Rows[i][j].ToString() + "\t");
                }
                sw.Write(Environment.NewLine);
            }

            sw.Flush();
            file.Flush();

            sw.Close();
            sw.Dispose();

            file.Close();
            file.Dispose();

            stopwatch.Stop();
            Console.WriteLine($"Save used {stopwatch.ElapsedMilliseconds} ms");
        }
    }
}
