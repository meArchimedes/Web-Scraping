using HtmlAgilityPack;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Threading.Tasks;
using System.Net;
using System.Text;
using System.IO;
namespace Web_Scraping
{
    public class Main
    {
        //Parses the URL and returns HtmlDocument object

        public static string url = "https://www.horeca.com/en/categorie/3621/combi-steamers";

        static HtmlDocument GetDocument(string url)
        {
            
        }
        public static HtmlDocument doc = GetDocument(url);
        HtmlNodeCollection linkNodes = doc.DocumentNode.SelectNodes("/html/body/div[3]/div[9]/div/div[2]/div[2]/div[3]/div");
        ///div/div/div/div/section/div[2]/ol/li[3]/article/div[2]/p[1]
        //*[@id="product-overview-container"]/div
        ///html/body/div[3]/div[9]/div/div[2]/div[2]/div[3]/div
    }
}

