using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelReader
{
    public class Article
    {
        public int Id { get; set; }
        public string Libelle { get; set; }
        public int PU { get; set; }

        public Article() { }

        public Article(string id, string libelle, string pu)
        {
            this.Id = int.Parse(id);
            this.Libelle = libelle;
            this.PU = int.Parse(pu);
        }
    }
}
