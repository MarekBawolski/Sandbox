using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Sandbox
{
    class Program
    {
        static void Main(string[] args)
        {
            //Read_From_Excel.getExcelFile();
            PoPllikach pp = new PoPllikach();
            List<string> pliki = pp.printFiles();

            foreach (string plik in pliki) {
                Console.WriteLine(plik);
            }


            foreach (string plik in pliki)
            {
                Read_From_Excel.getExcelFile(plik);
            }

            Console.ReadLine();

        }
    }
}
