using Microsoft.Office.Interop.Excel;
using System;
using _Excel = Microsoft.Office.Interop.Excel;
using System.Threading;
using System.Collections.Generic;


namespace excel_test_fuck
{
    class Program
    {
        static void Main(string[] args)
        {
            //開啟excel，等待RTD完整讀取資料，並隨意讀取一個值
            Excel excel = new Excel(@"D:\股票報價new", 1);
            excel.ReadCell(9,3);

            List<stock> stockslist = new List<stock>();
            List<stock> stockslist2 = new List<stock>();
            int i = 4;
            while (i <= 468)
            {
                try
                {
                    object[,] data2 = excel.Readrange(6, i, 12, i);
                    Console.WriteLine(data2[1, 1]);
                    Console.WriteLine(data2[2 ,1]);
                    Console.WriteLine(data2[ 3,1]);
                    Console.WriteLine(data2[4,1]);
                    Console.WriteLine(data2[5, 1]);
                    Console.WriteLine(data2[6, 1]);
                    Console.WriteLine(data2[7, 1]);
                    Console.WriteLine("成功"+i);
                    i++;
                }
                catch{ Console.WriteLine("失敗"); }
                Console.ReadLine();
            }
        }
        public class Excel
        {
            string path = "";
            _Application excel = new _Excel.Application();
            Workbook wb;
            Worksheet ws;
            public Excel(string path, int Sheet)
            {
                this.path = path;
                this.excel.Visible = true;
                wb = excel.Workbooks.Open(path, Type.Missing, true);
                ws = wb.Worksheets[Sheet];
            }
            public object[,] Readrange(int c, int d, int x, int y)
            {
               var a = ws.Range[ws.Cells[c, d], ws.Cells[x, y]].Value2;
               return a;
            }
            public void ReadCell(int a, int b)
            {
                Thread.Sleep(10000);
                double ccc = 0;
                while (ccc==0)
                {
                    try
                    { ccc = ws.Cells[a, b].Value2; }
                    catch { }
                }
            }
        }
        public class stock
        {
            public string name;
            public int number;
            public double price;
            public int volume;
            public int totalVolume;
            public double averange;
            public string dataTime;
        }
    }
}
