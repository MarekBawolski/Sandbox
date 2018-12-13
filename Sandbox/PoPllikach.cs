using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Sandbox
{
    class PoPllikach
    {
        private List<String> pliki;
        private List<string> ext;
        private string sciezka;

        public PoPllikach()
        {
            this.pliki = new List<string>();
            this.ext = new List<string> { ".xls", ".xlsx" };
            this.sciezka = @"C:\Users\Marek\Downloads\excel\input";


        }

        public List<string> printFiles()
        {

            var excelFiles = Directory.EnumerateFiles(sciezka, "*.xls", SearchOption.AllDirectories);
            foreach (string excel in excelFiles)
            {
                this.pliki.Add(excel);
            }
            return pliki;

        }

    }
}
