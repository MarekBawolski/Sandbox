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
        private string[] dir;
        private List<String> pliki;


        public PoPllikach()
        {
            this.dir =  Directory.GetFiles(@"C:\Users\Marek\Downloads\excel\input");
            this.pliki = new List<string>();

        }

        public List<string> printFiles() {
            foreach (string path in dir)
            {
                this.pliki.Add(path);
            }
            return pliki;
        }
    }
}
