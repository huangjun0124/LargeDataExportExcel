using System;

namespace ExportUsingOpenXML
{
    class Program
    {
        static void Main(string[] args)
        {
            TestOpenXML2();

            while (true)
            {
                Console.ReadLine();
                GC.Collect();
            }
        }

        static void TestOpenXML1()
        {
            // 内存占用依然很高，1.3 GB, 不过运行很快, 22 秒，生成的文件很小 2.11 MB
            OpenXMLTest ot = new OpenXMLTest();
            ot.DoTest();
        }

        static void TestOpenXML2()
        {
            // 内存占用低 190 MB，耗时少 18 秒，结束后内存占用 20MB， 生成的文件 26MB
            OpenXMLTest ot = new OpenXMLTest();
            ot.DoTest2();
        }
    }
}
