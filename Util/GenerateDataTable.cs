using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Util
{
    public class GenerateDataTable
    {
        public DataTable GetNewTable(int rowCount=600000)
        {
            Stopwatch sw = new Stopwatch();
            sw.Start();
            var data = new DataTable("测试表格");
            data.Columns.Add("Library", typeof(string));
            data.Columns.Add("Name", typeof(string));
            data.Columns.Add("Description", typeof(string));
            data.Columns.Add("Type", typeof(int));
            data.Columns.Add("Definer", typeof(string));
            data.Columns.Add("Definer_Description", typeof(string));
            data.Columns.Add("Creation_Date", typeof(DateTime));
            data.Columns.Add("Days_Since_Creation", typeof(string));
            data.Columns.Add("Size", typeof(decimal));
            data.Columns.Add("Last_Used", typeof(DateTime));
            data.Columns.Add("Attribute", typeof(string));
            data.Columns.Add("Count_Of_Objects_Referenced", typeof(string));
            data.Columns.Add("引用情况", typeof(bool));

            for (var i = 0; i < rowCount; i++)
            {
                DateTime? dt = null;
                if (i % 9 != 0)
                    dt = DateTime.Now;
                data.Rows.Add(
                    "xxxxxxxxx",
                    "xxxxxxxxx",
                    "错位符占用" + i,
                    i,
                    "xxxxxxxxx",
                    "xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx",
                    dt,
                    "789",
                    i + 1908.3456,
                    DateTime.Now.AddDays(120),
                    "GRC",
                    "0",
                    i % 3 == 0 ? true:false
                );
            }
            sw.Stop();
            Console.WriteLine($"Generate datable used [{sw.ElapsedMilliseconds}] ms");
            return data;
        }

        public DataTable GetTableWithNRows(DataTable dataIn, int rowCount)
        {
            DataTable data = dataIn.Clone();
            for (var i = 0; i < rowCount; i++)
            {
                data.Rows.Add(
                    "xxxxxxxxx",
                    "xxxxxxxxx",
                    "xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx",
                    "*USRSPC",
                    "xxxxxxxxx",
                    "xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx",
                    DateTime.Now,
                    "789",
                    "16384",
                    DateTime.Now.AddDays(120),
                    "GRC",
                    "0",
                    "0"
                );
            }

            return data;
        }
    }
}
