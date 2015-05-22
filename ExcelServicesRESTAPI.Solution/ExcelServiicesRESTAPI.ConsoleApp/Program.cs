using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Security;
using System.Text;
using System.Threading.Tasks;

namespace ExcelServiicesRESTAPI.ConsoleApp
{
    class Program
    {
        static void Main(string[] args)
        {
            //Conectando ao meu Tenant com SharePoint Online
            Console.WriteLine("Digite a url do seu SharePoint:");

            //Recuperando os dados do seu usuário
            var url = Console.ReadLine();
            Console.WriteLine("Digite seu usuário:");

            var usuario = Console.ReadLine();

            //Construindo os dados da sua senha
            Console.WriteLine("Digite sua senha:");
            var senha = new SecureString();

            var sair = false;

            while (true)
            {
                var chave = Console.ReadKey(true);

                switch (chave.Key)
                {
                    case ConsoleKey.Enter: sair = true; break;
                    case ConsoleKey.Escape: sair = true; return;
                    case ConsoleKey.Backspace:

                        if (senha.Length != 0)
                            senha.RemoveAt(senha.Length - 1);
                        break;
                    default: senha.AppendChar(chave.KeyChar); break;
                }

                if (sair) break;
            }

            senha.MakeReadOnly();

            //Criando as credenciais de acesso ao SharePoint
            var credenciais = new SharePointOnlineCredentials(usuario, senha);

            Console.WriteLine();
            Console.WriteLine("Iniciando requisição REST..................................");

            //Recuperando o nome da Biblioteca onde está seu arquivo
            Console.WriteLine();
            Console.WriteLine("Digite o nome da Biblioteca onde está seu arquivo Excel:");
            
            var biblioteca = Console.ReadLine();

            //Recuperando o nome do Arquivo Excel (não esqueça de colocar a extensão
            Console.WriteLine();
            Console.WriteLine("Digite o nome do seu arquivo Excel:");
            var nomeArquivo = Console.ReadLine();

            //Recuperando o tipo de solicitação desejado (Ranges, Tables, Charts, PivotTables)
            Console.WriteLine();
            Console.WriteLine("Tipo de Requisição:");
            var tipoRequisicao = Console.ReadLine();

            //Recuperando o tipo de solicitação desejado (Ranges, Tables, Charts, PivotTables)
            Console.WriteLine();
            Console.WriteLine("Dado que você deseja:");
            var dado= Console.ReadLine();

            //Recuperando o Formato do dado (atom, html, json)
            Console.WriteLine();
            Console.WriteLine("Formato do dado:");
            var formatodado = Console.ReadLine();

            url += String.Format("/_vti_bin/ExcelRest.aspx/{0}/{1}/Model/{2}('{3}')?$format={4}", biblioteca, nomeArquivo, tipoRequisicao, dado, formatodado);

            var requisicao = (HttpWebRequest)WebRequest.Create(url);
            requisicao.Credentials = credenciais;

            requisicao.Headers["X-FORMS_BASED_AUTH_ACCEPTED"] = "f";

            try
            {

                var resposta = (HttpWebResponse)requisicao.GetResponse();
                Stream responseStream = resposta.GetResponseStream();
                StreamReader responseReader = new StreamReader(responseStream);

                var resultado = responseReader.ReadToEnd();

                Console.WriteLine("Gerando arquivo com resultado..............................................");

                using (System.IO.StreamWriter file = new System.IO.StreamWriter(formatodado + ".txt", true))
                    file.WriteLine(resultado);

                Console.WriteLine("Arquivo gerado com sucesso!");
            }
            catch(Exception ex)
            {
                Console.WriteLine("Erro::" + ex.Message);
            }

            Console.ReadLine();
        }
    }
}
