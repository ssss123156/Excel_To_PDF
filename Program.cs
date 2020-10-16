using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace Excel_To_PDF
{
    class Program
    {
        static void Main(string[] args)
        {
            ExcelProcess excelProcess = excelProcess = new ExcelProcess(@"C:\Users\Alexander\Desktop\1.xlsx", @"C:\Users\Alexander\Desktop\2");
            Console.ReadKey();
        }
    }
}
