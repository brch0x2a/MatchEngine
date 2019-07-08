using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MatchEngineV00
{
    class Program
    {
        static void Main(string[] args)
        {
            string path = @"C:\Users\brch\Documents\UL\paquete\";
            string paros = "24_06_2019_08_46-paros"+".xls";
            string area = "test505" + ".xlsm";

            Matcher matcher = new Matcher();

            Console.WriteLine("-Match engine V00-");

            matcher.Path = path;
            matcher.ParosBook = path + paros;
            matcher.AreaBook = path + area;

            matcher.ReadParos();
            matcher.ReadArea();
            matcher.PreMatch();

            Console.ReadKey();
        }

    }
}
