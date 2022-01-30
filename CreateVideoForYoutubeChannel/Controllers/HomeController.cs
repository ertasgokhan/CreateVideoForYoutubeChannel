using CreateVideoForYoutubeChannel.Models;
using DocumentFormat.OpenXml.Packaging;
using HtmlAgilityPack;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;

namespace CreateVideoForYoutubeChannel.Controllers
{
    public class HomeController : Controller
    {
        private readonly ILogger<HomeController> _logger;

        public HomeController(ILogger<HomeController> logger)
        {
            _logger = logger;
        }

        public IActionResult Index()
        {
            var model = new YuotubeModel();

            return View(model);
        }

        [HttpPost]
        public IActionResult Index(YuotubeModel model)
        {
            // Epey

            string urlAddress = model.Url;
            string result = "";
            string docItems = "";

            HttpWebRequest httpRequest = (HttpWebRequest)HttpWebRequest.Create(urlAddress);
            httpRequest.Timeout = 10000;
            httpRequest.UserAgent = "Code Sample Web Client";
            httpRequest.Credentials = CredentialCache.DefaultCredentials;
            HttpWebResponse respone = (HttpWebResponse)httpRequest.GetResponse();
            StreamReader stream = new StreamReader(respone.GetResponseStream(), Encoding.UTF8);
            result = stream.ReadToEnd();


            string filepath = @"C:\Yutup\" + model.ProductName.Replace(" ", "").Trim() + ".txt";

            if (System.IO.File.Exists(filepath))
                System.IO.File.Delete(filepath);

            using (StreamWriter sw = System.IO.File.CreateText(filepath))
            {
                var doc = new HtmlAgilityPack.HtmlDocument();
                doc.LoadHtml(result);

                HtmlNode properties = doc.DocumentNode.SelectSingleNode("//div[@id='bilgiler']");

                if (properties != null)
                {
                    HtmlNodeCollection groups = doc.DocumentNode.SelectNodes(".//div[@id='grup']");

                    if (groups != null && groups.Count > 0)
                    {
                        string groupName = string.Empty;

                        for (int i = 0; i < groups.Count; i++)
                        {
                            HtmlNode newItem = groups[i];

                            groupName = newItem.SelectSingleNode(".//h3").InnerText;

                            docItems += string.Format("***{0}**** \n", groupName);

                            HtmlNode groupProperty = newItem.SelectSingleNode(".//ul[@class='grup']");

                            HtmlNodeCollection groupChilNodes = groupProperty.ChildNodes;

                            string propertyName = string.Empty;
                            string propertyValue = string.Empty;

                            foreach (var item2 in groupChilNodes)
                            {
                                if (item2.Name == "li")
                                {
                                    propertyName = item2.SelectSingleNode("strong").InnerText;
                                    propertyValue = item2.SelectSingleNode("span//a") != null ? item2.SelectSingleNode("span//a").InnerText.Replace("\n", "") : item2.SelectSingleNode("span//span").InnerText.Replace("\n", "");

                                    docItems += string.Format("{0}->{1} \n", propertyName, propertyValue);
                                }
                            }
                        }
                    }
                }

                sw.WriteLine(docItems);
            }

            // PPT

            List<string> phoneProperties = new List<string>();
            string filepath2 = @"C:\Yutup\SamsungGalaxyM22.txt";

            using (StreamReader rd = System.IO.File.OpenText(filepath2))
            {
                while (!rd.EndOfStream)
                {
                    string str = rd.ReadLine();
                    phoneProperties.Add(str);
                }
            }

            // just gets me the current location of the assembly to get a full path
            string fileName = @"C:\Yutup\Test1\VS_PPT - Modify 1.pptx";

            // open the presentation in edit mode -> the bool parameter stands for 'isEditable'
            using (PresentationDocument document = PresentationDocument.Open(fileName, true))
            {
                // going through the slides of the presentation
                foreach (SlidePart slidePart in document.PresentationPart.SlideParts)
                {
                    // searching for a text with the placeholder i want to replace
                    DocumentFormat.OpenXml.Drawing.Text text =
                        slidePart.RootElement.Descendants<DocumentFormat.OpenXml.Drawing.Text>().FirstOrDefault(x => x.Text == "Product1_Çıkış_Yılı");

                    // change the text
                    if (text != null)
                        text.Text = phoneProperties.FirstOrDefault(x => x.Contains("Çıkış Tarihi")).Split("->")[1]; 

                    // searching for the second text with the placeholder i want to replace
                    text =
                        slidePart.RootElement.Descendants<DocumentFormat.OpenXml.Drawing.Text>().FirstOrDefault(x => x.Text == "Product2_Çıkış_Yılı");

                    // change the text
                    if (text != null)
                        text.Text = phoneProperties.FirstOrDefault(x => x.Contains("Çıkış Tarihi")).Split("->")[1];
                }

                document.Save();
            }

            return View(model);
        }

        public IActionResult Privacy()
        {
            return View();
        }

        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error()
        {
            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }
    }
}
