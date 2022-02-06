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

            List<string> phoneProperties = new List<string>();
            string filepath2 = @"C:\Yutup\PHONE\" + phoneVSModel.Phone1.Replace(" ", "") + ".txt";

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

            return View();
        }
    }
}
