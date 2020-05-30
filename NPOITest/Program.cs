using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace NPOITest
{
    class Program
    {
        static void Main(string[] args)
        {
            // 内存占用【2600MB】,耗时【50 秒】,生成的文件大小【29 MB】
            ExportUsingNPOI test = new ExportUsingNPOI();
            test.DoTest();

            while (true)
            {
                Console.ReadLine();
                GC.Collect();
            }
        }
    }
}
