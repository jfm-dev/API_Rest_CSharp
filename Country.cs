using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CREST
{
    public class Country
    {
        public string Nome { get; set; }
        public string Capital { get; set; }
        public string? Regiao { get; set; }
        public string? Subregiao { get; set; }
        public long Populacao { get; set; }
        public string Area { get; set; }
        public string? Fusohorario { get; set; }
        public string? Nomenativo { get; set; }
        public string? Bandeira { get; set; }
    }
}
