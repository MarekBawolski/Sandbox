using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using CsvHelper;

namespace Sandbox
{
    class Program
    {
        static void Main(string[] args)
        {
            PoPllikach pp = new PoPllikach();
            List<string> pliki = pp.printFiles();
            StreamWriter writer = new StreamWriter(@"C:\Users\Marek\Downloads\excel\out\log.txt", true);

            foreach (string plik in pliki)
            {
                MatchCollection matches = Regex.Matches(plik, "POL.........");
                List<string> chemia = Read_From_Excel.getExcelFile(plik, writer);



                foreach (Match ma in matches)
                {

                    if (chemia.Count<2) {
                        continue;
                    }
                    string content = ma +";";
                    string currentPlik = plik;
                    foreach (string chemikal in chemia)
                    {
                        if (chemikal == null)
                        {
                            content += "null";
                        }
                        content += chemikal+";";


                    }
                    using (StreamWriter w = File.AppendText(@"C:\Users\Marek\Downloads\excel\out\output.csv"))
                    {
                        w.WriteLine(content);
                    }
                }
                


            }

        }
    }
}
