using System;
using System.Collections.Generic;
using System.Text;

namespace WebScrapper.Model
{
    public class Empresa
    {
        public Empresa()
        {
            this.Telefones = new List<Telefone>();
        }
        public string Categoria { get; set; }
        public string Latitude { get; set; }
        public string Longitude { get; set; }
        public string Nome { get; set; }
        public string Endereco { get; set; }
        public string Cidade { get; set; }
        public string Estado { get; set; }
        public string CEP { get; set; }
        public List<Telefone> Telefones { get; set; }
    }
}
