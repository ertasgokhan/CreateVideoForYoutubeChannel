using CreateVideoForYoutubeChannel.Models;
using DocumentFormat.OpenXml.Packaging;
using HtmlAgilityPack;
using Microsoft.AspNetCore.Mvc;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;

namespace CreateVideoForYoutubeChannel.Controllers
{
    public class RobotVacuumController : Controller
    {
        public IActionResult Index()
        {
            var model = new YouTubeModel();

            return View(model);
        }

        [HttpPost]
        public IActionResult Index(YouTubeModel model)
        {
            string result = "";
            string docItems = "";

            HttpWebRequest httpRequest = (HttpWebRequest)HttpWebRequest.Create(model.Url);
            httpRequest.Timeout = 10000;
            httpRequest.UserAgent = "Code Sample Web Client";
            httpRequest.Credentials = CredentialCache.DefaultCredentials;
            HttpWebResponse respone = (HttpWebResponse)httpRequest.GetResponse();
            StreamReader stream = new StreamReader(respone.GetResponseStream(), Encoding.UTF8);
            result = stream.ReadToEnd();

            var doc = new HtmlAgilityPack.HtmlDocument();
            doc.LoadHtml(result);

            HtmlNode productName = doc.DocumentNode.SelectSingleNode("//div[@class='baslik']//h1//a");
            model.ProductName = productName.InnerText;

            string filepath = @"C:\Yutup\ROBOTVACUUM\" + model.ProductName.Replace(" ", "").Replace("/", "").Trim() + ".txt";

            if (System.IO.File.Exists(filepath))
                System.IO.File.Delete(filepath);

            using (StreamWriter sw = System.IO.File.CreateText(filepath))
            {
                docItems += string.Format("Ürün->{0} \n", model.ProductName);
                docItems += string.Format("Url->{0} \n", model.Url);

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

                            foreach (var item2 in groupChilNodes)
                            {
                                string propertyName = string.Empty;
                                string propertyValue = string.Empty;

                                if (item2.Name == "li")
                                {
                                    propertyName = item2.SelectSingleNode("strong").InnerText;

                                    if ((item2.SelectNodes("span/span").Count < 2 && (item2.SelectSingleNode("span//a") != null || item2.SelectSingleNode("span//span") != null)) || item2.SelectSingleNode("span//span[@class='degerYok']") != null || item2.SelectSingleNode("span//span[@class='degerVar']") != null)
                                    {
                                        propertyValue = item2.SelectSingleNode("span//span").InnerText.Replace("\n", "");
                                    }
                                    else
                                    {
                                        HtmlNodeCollection propertyGroups = item2.SelectNodes(".//span[@class='']");

                                        if (propertyGroups != null)
                                        {
                                            foreach (var item in propertyGroups)
                                            {
                                                if (string.IsNullOrEmpty(propertyValue))
                                                {
                                                    propertyValue = item.InnerText.Replace("\n", "");
                                                }
                                                else
                                                {
                                                    propertyValue += ", " + item.InnerText.Replace("\n", "");
                                                }
                                            }
                                        }
                                    }

                                    docItems += string.Format("{0}->{1} \n", propertyName, propertyValue);
                                }
                            }
                        }
                    }
                }

                sw.WriteLine(docItems);
            }

            ViewData["Message"] = "Dosya Başarıyla Oluşturulmuştur";

            return View(model);
        }

        public IActionResult RobotVacuumVS()
        {
            return View();
        }

        [HttpPost]
        public IActionResult RobotVacuumVS(RobotVacuumModel robotVacuumModel)
        {
            // PPT

            List<string> robotVacuum1Properties = new List<string>();
            List<string> robotVacuum2Properties = new List<string>();

            string filepathPhone1 = @"C:\Yutup\ROBOTVACUUM\" + robotVacuumModel.RobotVacuum1.Replace(" ", "") + ".txt";
            string filepathPhone2 = @"C:\Yutup\ROBOTVACUUM\" + robotVacuumModel.RobotVacuum2.Replace(" ", "") + ".txt";

            using (StreamReader rd = System.IO.File.OpenText(filepathPhone1))
            {
                while (!rd.EndOfStream)
                {
                    string str = rd.ReadLine();
                    robotVacuum1Properties.Add(str);
                }
            }

            using (StreamReader rd = System.IO.File.OpenText(filepathPhone2))
            {
                while (!rd.EndOfStream)
                {
                    string str = rd.ReadLine();
                    robotVacuum2Properties.Add(str);
                }
            }

            // just gets me the current location of the assembly to get a full path
            string fileName = @"C:\Yutup\ROBOTVACUUM\VS\VS_PPT_ROBOTVACUUM.pptx";

            // open the presentation in edit mode -> the bool parameter stands for 'isEditable'
            using (PresentationDocument document = PresentationDocument.Open(fileName, true))
            {
                string robotVacuum1NewItem = string.Empty;
                string robotVacuum2NewItem = string.Empty;

                foreach (var item in robotVacuum1Properties)
                {
                    robotVacuum1NewItem = string.Format("Product1_{0}", item.Split("->")[0].Replace(" ", "_").Replace("/", "").Replace("(", "").Replace(")", "").Replace(".", ""));

                    // going through the slides of the presentation
                    foreach (SlidePart slidePart in document.PresentationPart.SlideParts)
                    {
                        // searching for a text with the placeholder i want to replace
                        DocumentFormat.OpenXml.Drawing.Text text =
                            slidePart.RootElement.Descendants<DocumentFormat.OpenXml.Drawing.Text>().FirstOrDefault(x => x.Text == robotVacuum1NewItem);

                        // change the text
                        if (text != null)
                            text.Text = robotVacuum1Properties.FirstOrDefault(x => x.Contains(item)).Split("->")[1].Replace("Var", "Yes").Replace("Yok", "No").Replace("Milyon", "Million").Replace("Milyar", "Billion").Replace("Cam", "Glass").Replace("Alüminyum", "Aluminum").Replace("Gram", "Grams").Replace("Çekirdek", "Core").Replace("Kablosu", "Cable").Replace("Piksel", "Pixel").Replace("Çift Hat", "Dual SIM").Replace("Ocak", "January").Replace("Şubat", "February").Replace("Mart", "March").Replace("Nisan", "April").Replace("Mayıs", "May").Replace("Haziran", "June").Replace("Temmuz", "July").Replace("Ağustos", "August").Replace("Eylül", "September").Replace("Ekim", "October").Replace("Kasım", "November").Replace("Aralık", "December").Replace("Robot Süpürge", "Robot Vacuum").Replace("Hard Floor", "Sert Zemin").Replace("Halı", "Carpet").Replace("saat", "hour(s)").Replace("dk", "min").Replace("Süpürme", "Vacuum").Replace("Mobil Uygulama", "Mobile Application").Replace("Adet", "piece(s)").Replace("Siyah", "Black").Replace("Beyaz", "White").Replace("Mobil Uygulama İle Otomatik", "By Mobile Application").Replace("Google Asistan", "Google Assistant").Replace("Lazer", "Laser");
                    }
                }

                foreach (var item in robotVacuum2Properties)
                {
                    robotVacuum2NewItem = string.Format("Product2_{0}", item.Split("->")[0].Replace(" ", "_").Replace("/", "").Replace("(", "").Replace(")", "").Replace(".", ""));

                    // going through the slides of the presentation
                    foreach (SlidePart slidePart in document.PresentationPart.SlideParts)
                    {
                        // searching for a text with the placeholder i want to replace
                        DocumentFormat.OpenXml.Drawing.Text text =
                            slidePart.RootElement.Descendants<DocumentFormat.OpenXml.Drawing.Text>().FirstOrDefault(x => x.Text == robotVacuum2NewItem);

                        // change the text
                        if (text != null)
                            text.Text = robotVacuum2Properties.FirstOrDefault(x => x.Contains(item)).Split("->")[1].Replace("Var", "Yes").Replace("Yok", "No").Replace("Milyon", "Million").Replace("Milyar", "Billion").Replace("Cam", "Glass").Replace("Alüminyum", "Aluminum").Replace("Gram", "Grams").Replace("Çekirdek", "Core").Replace("Kablosu", "Cable").Replace("Piksel", "Pixel").Replace("Çift Hat", "Dual SIM").Replace("Ocak", "January").Replace("Şubat", "February").Replace("Mart", "March").Replace("Nisan", "April").Replace("Mayıs", "May").Replace("Haziran", "June").Replace("Temmuz", "July").Replace("Ağustos", "August").Replace("Eylül", "September").Replace("Ekim", "October").Replace("Kasım", "November").Replace("Aralık", "December").Replace("Robot Süpürge", "Robot Vacuum").Replace("Hard Floor", "Sert Zemin").Replace("Halı", "Carpet").Replace("saat", "hour(s)").Replace("dk", "min").Replace("Süpürme", "Vacuum").Replace("Mobil Uygulama", "Mobile Application").Replace("Adet", "piece(s)").Replace("Siyah", "Black").Replace("Beyaz", "White").Replace("Mobil Uygulama İle Otomatik", "By Mobile Application").Replace("Google Asistan", "Google Assistant").Replace("Lazer", "Laser");
                    }
                }

                document.Save();
            }

            ViewData["Message"] = "PPT Başarıyla Güncellenmiştir";

            return View(robotVacuumModel);
        }
    }
}
