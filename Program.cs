using System.Net;
using Microsoft.Office.Interop.Word;
using Xceed.Words.NET;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;


namespace RPA_Json_IdWall
{
    class Program
    {
        public static Microsoft.Office.Interop.Word.Document WordDocument { get; set; }

        public static void Main(string[] args)
        {
            string CaminhoArquivo = Directory.GetCurrentDirectory() + @"\ExData\Modelo_Relatorio_Idwall.docx";
            string jReturn = GetResponse();
            var jObj = (JObject)JsonConvert.DeserializeObject(jReturn);

            //Console.WriteLine(jReturn); //Printa o retorno JSON

            var nomeCliente =            jObj["result"]["cpf"]["nome"].ToString();
            var cpfCliente =             jObj["result"]["cpf"]["numero"].ToString();
            var status =                 jObj["result"]["mensagem"].ToString();
            var dataConsulta =           jObj["result"]["criado_em"].ToString();
            var codigoRelatorio =        jObj["result"]["numero"].ToString();
            var dataNascimento =         jObj["result"]["cpf"]["data_de_nascimento"].ToString();
            var dataInscricao =          jObj["result"]["cpf"]["cpf_data_de_inscricao"].ToString();
            var situacaoCadastral =      jObj["result"]["cpf"]["cpf_situacao_cadastral"].ToString();
            var resultDoc =              jObj["result"]["resultado"].ToString();
            var detalheDocumentoscopia = jObj["result"]["documentoscopia"]["result"]["note"].ToString();
            var valorDocumentoscopia =   jObj["result"]["documentoscopia"]["result"]["value"].ToString();
            var rgChancela =             jObj["result"]["documentoscopia"]["evidence"][0]["value"].ToString();
            var rgFonte =                jObj["result"]["documentoscopia"]["evidence"][1]["value"].ToString();
            var rgAlinhamento =          jObj["result"]["documentoscopia"]["evidence"][2]["value"].ToString();
            var rgPerfuracao =           jObj["result"]["documentoscopia"]["evidence"][3]["value"].ToString();
            var rgBrasao =               jObj["result"]["documentoscopia"]["evidence"][4]["value"].ToString();
            var rgAnalise =              jObj["result"]["documentoscopia"]["evidence"][5]["value"].ToString();
            var rgOrgao =                jObj["result"]["documentoscopia"]["evidence"][6]["value"].ToString();
            

            using (var doc = DocX.Load(CaminhoArquivo))
            {
                doc.ReplaceText("#nomeCliente", nomeCliente);
                doc.ReplaceText("#cpfCliente", cpfCliente);
                doc.ReplaceText("#status", status);
                doc.ReplaceText("#dataConsulta", dataConsulta);
                doc.ReplaceText("#codigoRelatorio", codigoRelatorio);
                doc.ReplaceText("#dataNascimento", dataNascimento);
                doc.ReplaceText("#dataInscricao", dataInscricao);
                doc.ReplaceText("#situacaoCadastral", situacaoCadastral);
                doc.ReplaceText("#resultDoc", resultDoc);
                doc.ReplaceText("#detalheDocumentoscopia", detalheDocumentoscopia);
                doc.ReplaceText("#valorDocumentoscopia", valorDocumentoscopia);
                doc.ReplaceText("#rgChancela", rgChancela);
                doc.ReplaceText("#rgFonte", rgFonte);
                doc.ReplaceText("#rgAlinhamento", rgAlinhamento);
                doc.ReplaceText("#rgPerfuracao", rgPerfuracao);
                doc.ReplaceText("#rgBrasao", rgBrasao);
                doc.ReplaceText("#rgAnalise", rgAnalise);
                doc.ReplaceText("#rgOrgao", rgOrgao);

                doc.Save();
            }

            ConverterParaPDF(CaminhoArquivo);

            Console.WriteLine("Trasnformação concluida.");
        }
        
        static string GetResponse()
        {
            //Chave token
            string key = "";

            //Instancia a requisição
            ServicePointManager.Expect100Continue = false;
            var request = (HttpWebRequest)WebRequest.Create("https://api-v2.idwall.co/relatorios/e5d070c2-57eb-4f4f-a9a2-a0327538ace4/dados");

            //Seleciona o método GET
            request.Method = "GET";

            //Configura o Header
            request.Headers.Add("Authorization", key);
            request.Headers.Add("Content-Type", "application/json");

            var ret = (HttpWebResponse)request.GetResponse();
            var response = new StreamReader(ret.GetResponseStream()).ReadToEnd();

            return response;
        }
        
        public static void ConverterParaPDF(string origem)
        {
            string pdfSaida = origem.Substring(0, origem.Length - 4) + "pdf";

            Application appWord = new Application();
            WordDocument = appWord.Documents.Open(origem);
            WordDocument.ExportAsFixedFormat(pdfSaida, WdExportFormat.wdExportFormatPDF);
            WordDocument.Close();
            appWord.Quit();
            File.Delete(origem);
        }
    }

}
