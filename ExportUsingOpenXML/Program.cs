using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExportUsingClosedXML
{
    class Program
    {
        static void Main(string[] args)
        {
            TestClosedXML1();
            while (true)
            {
                Console.ReadLine();
                GC.Collect();
            }
        }

        static void TestUsingTabDelimi()
        {
            // 内存占用 180MB，耗时 3秒 ，但是文件打不开，格式不对；文件过大【110 MB】
            ExportUsingTabDelimi etd = new ExportUsingTabDelimi();
            etd.DoTest();
        }

        static void TestClosedXML2()
        {
            // 运行速度很慢，内存占用越往后也越大
            ClosedXmlTest ct = new ClosedXmlTest();
            ct.DoTest_GenerateThenAttach();
            ct = null;
        }

        static void TestClosedXML1()
        {
            // Memory usage will exceed 3000 MB, drop quickly to 60 MB after gc done
            // 运行时间 100 秒，文件大小 23 MB
            ClosedXmlTest ct = new ClosedXmlTest();
            ct.DoTest_GenerateDirectly();
            ct = null;
        }
    }
}
