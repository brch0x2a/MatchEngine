using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MatchEngineV00
{
    class Paro
    {
        public string linea { get; set; }
        public int dia { get; set; }
        public int turno { get; set; }
        public int codigo { get; set; }
        public int minutos { get; set; }

        override
        public string ToString() {
            return linea + "|\t" + codigo + "|\t" + dia + "\t" + turno + "\t" + minutos;
        }
    }
}
