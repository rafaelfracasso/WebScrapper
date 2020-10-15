using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Net;
using HtmlAgilityPack;
using WebScrapper.Model;

namespace WebScrapper
{
    class Program
    {
        static void Main(string[] args)
        {
            string url = "https://www.guiamais.com.br/encontre?searchbox=true&what=&where=Cuiab%C3%A1%2C+MT&order=alpha";

            HtmlWeb web = new HtmlWeb();
            HtmlDocument doc = web.Load(url);

            List<Empresa> empresas = new List<Empresa>();

            var list = doc.DocumentNode.SelectNodes("//meta[@itemprop= 'url']");
            foreach (var node in list)
            {
                Empresa empresa = new Empresa();

                string contentUrl = node.GetAttributeValue("content", "");

                HtmlDocument contentDocument = web.Load(contentUrl);
                var inputCategoria = contentDocument.DocumentNode.SelectSingleNode("//input[@type='hidden' and @id='tx_category']").Attributes["value"];
                var inputCidade = contentDocument.DocumentNode.SelectSingleNode("//input[@type='hidden' and @id='tx_city']").Attributes["value"];
                var inputLatitude = contentDocument.DocumentNode.SelectSingleNode("//input[@type='hidden' and @id='tx_lat']").Attributes["value"];
                var inputLongitude = contentDocument.DocumentNode.SelectSingleNode("//input[@type='hidden' and @id='tx_lng']").Attributes["value"];

                if (inputCategoria.Value != null)
                {
                    empresa.Categoria = inputCategoria.Value;
                }

                if (inputCidade.Value != null)
                {
                    empresa.Cidade = inputCidade.Value;
                }

                if (inputLatitude.Value != null)
                {
                    empresa.Latitude = inputLatitude.Value;
                }

                if (inputLongitude.Value != null)
                {
                    empresa.Longitude = inputLongitude.Value;
                }

                empresa.Nome = contentDocument.DocumentNode.SelectSingleNode("//h1[@class='tp-companyName']").InnerText;
                empresa.Endereco = contentDocument.DocumentNode.SelectSingleNode("//span[@class='tp-address']").InnerText;
                empresa.Estado = contentDocument.DocumentNode.SelectSingleNode("//span[@class='tp-state']").InnerText;
                empresa.CEP = contentDocument.DocumentNode.SelectSingleNode("//span[@class='tp-postalCode']").InnerText;

                var telefoneList = contentDocument.DocumentNode.SelectNodes("//p[@class='phone detail tp-phone']");
                if (telefoneList != null)
                {
                    foreach (var telefoneValue in telefoneList)
                    {
                        Telefone telefone = new Telefone();
                        telefone.Numero = telefoneValue.InnerText;
                        empresa.Telefones.Add(telefone);
                    }
                }

                empresas.Add(empresa);
            }

            DataTable dt = CreateDataTable(empresas);

            XLWorkbook wb = new XLWorkbook();
            wb.Worksheets.Add(dt, "WorksheetName");
            wb.SaveAs(@"d:\clientes.xlsx", false);
        }


        private static DataTable CreateDataTable(List<Empresa> empresas)
        {
            DataTable dt = new DataTable();
            dt.Columns.Add(new DataColumn("Categoria"));
            dt.Columns.Add(new DataColumn("Latitude"));
            dt.Columns.Add(new DataColumn("Longitude"));
            dt.Columns.Add(new DataColumn("Nome"));
            dt.Columns.Add(new DataColumn("Endereco"));
            dt.Columns.Add(new DataColumn("Cidade"));
            dt.Columns.Add(new DataColumn("Estado"));
            dt.Columns.Add(new DataColumn("CEP"));
            int numeroDeTelefones = empresas.Max(e => e.Telefones.Count());
            for (int i = 0; i < numeroDeTelefones; i++)
            {
                dt.Columns.Add(new DataColumn("Telefone " + (i+1).ToString()));
            }

            foreach (Empresa empresa in empresas)
            {
                DataRow newRow = dt.NewRow();

                newRow["Categoria"] = empresa.Categoria;
                newRow["Latitude"] = empresa.Latitude;
                newRow["Longitude"] = empresa.Longitude;
                newRow["Nome"] = empresa.Nome;
                newRow["Endereco"] = empresa.Endereco;
                newRow["Cidade"] = empresa.Cidade;
                newRow["Estado"] = empresa.Estado;
                newRow["CEP"] = empresa.CEP;
                for (int i = 0; i < empresa.Telefones.Count(); i++)
                {
                    newRow["Telefone " + (i + 1).ToString()] = empresa.Telefones[i].Numero;
                }
                dt.Rows.Add(newRow);
            }

            return dt;
        }
    }
}
