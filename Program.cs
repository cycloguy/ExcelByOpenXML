using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelByOpenXML
{
    class Program
    {
        static void Main(string[] args)
        {
            Report report = new Report();
            report.CreateExcelDoc(@"D:\Report.xlsx");
            Console.ReadKey();
        }
    }
}
