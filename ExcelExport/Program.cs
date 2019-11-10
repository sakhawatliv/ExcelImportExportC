using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using _Excel = Microsoft.Office.Interop.Excel;


namespace ExcelExport
{
    class Program
    {
        static void Main(string[] args)
        {
            Exel excel = new Exel(@"C:\Test.xlsx", 1);
            Console.WriteLine(excel.ReadCell(0, 0));
            Console.ReadKey();
        }
    }
}
