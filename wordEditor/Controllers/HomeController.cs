using System;
using System.IO;
using System.Web.Mvc;
using DocumentFormat.OpenXml.Packaging;
using OpenXmlPowerTools;
using System.Xml.Linq;
using System.Web;
using System.Drawing;
using System.Drawing.Imaging;

public class HomeController : Controller
{
    public ActionResult Index()
    {
        return View();
    }

    [HttpPost]
    public ActionResult Upload(HttpPostedFileBase file)
    {
        if (file != null && Path.GetExtension(file.FileName) == ".docx")
        {
            try
            {
                // Dosyayı sunucuya kaydet
                string fileName = Path.GetFileName(file.FileName);
                string filePath = Server.MapPath("~/UploadedFiles/") + fileName;
                file.SaveAs(filePath);

                // .docx dosyasından metin içeriğini çıkar
                string htmlContent = ConvertDocToHtml(filePath, Server.MapPath("~/UploadedFiles/") + "output.html", this);

                // HTML içeriğini ViewBag'e ekle
                ViewBag.HtmlContent = htmlContent;
                return View("Index");
            }
            catch (Exception ex)
            {
                // Hataları ele al (örneğin, dosya bulunamadı, çıkarma hatası)
                ViewBag.ErrorMessage = "Error: " + ex.Message;
                return View("Index");
            }
        }
        else
        {
            ViewBag.ErrorMessage = "Please select a valid .docx file.";
            return View("Index");
        }
    }


    public static string ConvertDocToHtml(string sourcePath, string targetPath, Controller controller)
    {
        byte[] byteArray = System.IO.File.ReadAllBytes(sourcePath);
        using (MemoryStream memoryStream = new MemoryStream())
        {
            memoryStream.Write(byteArray, 0, byteArray.Length);
            using (WordprocessingDocument doc = WordprocessingDocument.Open(memoryStream, true))
            {
                HtmlConverterSettings settings = new HtmlConverterSettings()
                {
                    PageTitle = "My Page Title"
                };

                settings.ImageHandler = imageInfo =>
                {
                    try
                    {

                        string imageFileName = "image" + Guid.NewGuid().ToString("N") + ".jpg";
                        string imageFilePath = controller.Server.MapPath("~/UploadedFiles/Images/") + imageFileName;


                        imageInfo.Bitmap.Save(imageFilePath, ImageFormat.Jpeg);


                        XElement img = new XElement("img",
                            new XAttribute("src", "/UploadedFiles/Images/" + imageFileName),
                            imageInfo.ImgStyleAttribute,
                            imageInfo.AltText != null ? new XAttribute("alt", imageInfo.AltText) : null);

                        return img;
                    }
                    catch (System.Runtime.InteropServices.ExternalException)
                    {
                        return null;
                    }
                };

                XElement html = HtmlConverter.ConvertToHtml(doc, settings);
                
                string imagesFolderPath = controller.Server.MapPath("~/UploadedFiles/Images/");
                if (!Directory.Exists(imagesFolderPath))
                {
                    Directory.CreateDirectory(imagesFolderPath);
                }

                System.IO.File.WriteAllText(targetPath, html.ToStringNewLineOnAttributes());
                return html.ToString();
            }
        }
    }

}
