using CreateVideoForYoutubeChannel.Models;
using DocumentFormat.OpenXml.Packaging;
using Microsoft.AspNetCore.Mvc;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

namespace CreateVideoForYoutubeChannel.Controllers
{
    public class PhoneController : Controller
    {
        public IActionResult Index()
        {
            return View();
        }

        public IActionResult PhoneVS()
        {

            return View();
        }

        [HttpPost]
        public IActionResult PhoneVS(PhoneVSModel phoneVSModel)
        {
            // PPT

            List<string> phone1Properties = new List<string>();
            List<string> phone2Properties = new List<string>();

            string filepathPhone1 = @"C:\Yutup\PHONE\" + phoneVSModel.Phone1.Replace(" ", "") + ".txt";
            string filepathPhone2 = @"C:\Yutup\PHONE\" + phoneVSModel.Phone2.Replace(" ", "") + ".txt";

            using (StreamReader rd = System.IO.File.OpenText(filepathPhone1))
            {
                while (!rd.EndOfStream)
                {
                    string str = rd.ReadLine();
                    phone1Properties.Add(str);
                }
            }

            using (StreamReader rd = System.IO.File.OpenText(filepathPhone2))
            {
                while (!rd.EndOfStream)
                {
                    string str = rd.ReadLine();
                    phone2Properties.Add(str);
                }
            }

            // just gets me the current location of the assembly to get a full path
            string fileName = @"C:\Yutup\PHONE\VS\VS_PPT_PHONE.pptx";

            // open the presentation in edit mode -> the bool parameter stands for 'isEditable'
            using (PresentationDocument document = PresentationDocument.Open(fileName, true))
            {
                string phone1NewItem = string.Empty;
                string phone2NewItem = string.Empty;

                foreach (var item in phone1Properties)
                {
                    phone1NewItem = string.Format("Product1_{0}", item.Split("->")[0].Replace(" ", "_"));

                    // going through the slides of the presentation
                    foreach (SlidePart slidePart in document.PresentationPart.SlideParts)
                    {
                        // searching for a text with the placeholder i want to replace
                        DocumentFormat.OpenXml.Drawing.Text text =
                            slidePart.RootElement.Descendants<DocumentFormat.OpenXml.Drawing.Text>().FirstOrDefault(x => x.Text == phone1NewItem);

                        // change the text
                        if (text != null)
                            text.Text = phone1Properties.FirstOrDefault(x => x.Contains(item)).Split("->")[1];
                    }
                }

                foreach (var item in phone2Properties)
                {
                    phone2NewItem = string.Format("Product2_{0}", item.Split("->")[0].Replace(" ", "_"));

                    // going through the slides of the presentation
                    foreach (SlidePart slidePart in document.PresentationPart.SlideParts)
                    {
                        // searching for a text with the placeholder i want to replace
                        DocumentFormat.OpenXml.Drawing.Text text =
                            slidePart.RootElement.Descendants<DocumentFormat.OpenXml.Drawing.Text>().FirstOrDefault(x => x.Text == phone2NewItem);

                        // change the text
                        if (text != null)
                            text.Text = phone2Properties.FirstOrDefault(x => x.Contains(item)).Split("->")[1];
                    }
                }

                //// going through the slides of the presentation
                //foreach (SlidePart slidePart in document.PresentationPart.SlideParts)
                //{
                //    // searching for the second text with the placeholder i want to replace
                //    text =
                //        slidePart.RootElement.Descendants<DocumentFormat.OpenXml.Drawing.Text>().FirstOrDefault(x => x.Text == "Product2_Çıkış_Yılı");

                //    // change the text
                //    if (text != null)
                //        text.Text = phone1Properties.FirstOrDefault(x => x.Contains("Çıkış Tarihi")).Split("->")[1];
                //}

                document.Save();
            }

            ViewData["Message"] = "PPT Başarıyla Güncellenmiştir";

            return View(phoneVSModel);
        }
    }
}
