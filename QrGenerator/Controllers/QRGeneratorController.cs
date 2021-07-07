using Microsoft.AspNetCore.Mvc;
using QRCoder;
using System.IO;
using DocumentFormat.OpenXml.Packaging;
using System;
using DocumentFormat.OpenXml.Wordprocessing;
// using System.Linq;
using iTextSharp.text;
using iTextSharp.text.pdf;
using System.Collections.Generic;
using QrGenerator.Attributes;

namespace QrGenerator.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    [ApiKey]
    public class QRGeneratorController : ControllerBase
    {
        public class DocumentQR
        {
            public string url { get; set; }
            public string watermark { get; set; }
            public string base64 { get; set; }
            public string placement { get; set; }
        }

        public class Result
        {
            public string contentType { get; set; }
            public string content { get; set; }

            public Result(string contentTypeName, string contentData)
            {
                contentType = contentTypeName;
                content = contentData;
            }
        }

        //Додає QR у pdf документ
        [HttpPost("pdf")]
        public ActionResult<object> AddQRToPDF([FromBody] DocumentQR pdf)
        {
            #region Генерація QR коду
            QRCodeGenerator qrGenerator = new QRCodeGenerator();

            QRCodeData qrCodeData = qrGenerator.CreateQrCode(Helpers.TinyLink(pdf.url), QRCodeGenerator.ECCLevel.L);
            PngByteQRCode qrCode = new PngByteQRCode(qrCodeData);
            byte[] qrCodeAsPngByteArr = qrCode.GetGraphic(20);
            #endregion

            #region Вставка QR коду у pdf документ
            try
            {
                var bytes = Convert.FromBase64String(pdf.base64);
                MemoryStream _pdfMem = new MemoryStream(0);
                _pdfMem.Write(bytes, 0, bytes.Length);

                var reader = new PdfReader(bytes);
                MemoryStream _output = new MemoryStream(0);
                var stamper = new PdfStamper(reader, _output);

                if (pdf.placement == "last")
                {
                    var pdfContent = stamper.GetOverContent(reader.NumberOfPages);

                    using (MemoryStream qrStream = new MemoryStream(qrCodeAsPngByteArr))
                    {
                        Image img = Image.GetInstance(qrStream);

                        var size = reader.GetPageSize(reader.NumberOfPages);

                        img.SetAbsolutePosition(size.Width - 50, 20);
                        img.ScaleAbsoluteHeight(35);
                        img.ScaleAbsoluteWidth(35);

                        pdfContent.AddImage(img);
                    }
                }
                else
                if (pdf.placement == "first")
                {
                    var pdfContent = stamper.GetOverContent(1);

                    using (MemoryStream qrStream = new MemoryStream(qrCodeAsPngByteArr))
                    {
                        Image img = Image.GetInstance(qrStream);

                        var size = reader.GetPageSize(reader.NumberOfPages);

                        img.SetAbsolutePosition(size.Width - 50, 20);
                        img.ScaleAbsoluteHeight(35);
                        img.ScaleAbsoluteWidth(35);

                        pdfContent.AddImage(img);
                    }

                }
                else
                {
                    for (int i = 1; i <= reader.NumberOfPages; i++)
                    {
                        var pdfContent = stamper.GetOverContent(i);

                        using (MemoryStream qrStream = new MemoryStream(qrCodeAsPngByteArr))
                        {
                            Image img = Image.GetInstance(qrStream);

                            var size = reader.GetPageSize(reader.NumberOfPages);

                            img.SetAbsolutePosition(size.Width - 50, 20);
                            img.ScaleAbsoluteHeight(35);
                            img.ScaleAbsoluteWidth(35);

                            pdfContent.AddImage(img);

                            Helpers.AddWaterMark(pdf.watermark, reader, stamper);
                        }
                    }
                }

                stamper.Close();
                reader.Close();

                return new Result("application/pdf", Convert.ToBase64String(_output.ToArray()));
            }
            catch (Exception ex)
            {
                return ex;
            }
            #endregion
        }

        // Post api/qrgenerator Додає QR у docx документ
        [HttpPost]
        public ActionResult<object> AddQRToDocument([FromBody] DocumentQR document)
        {
            #region Генерація QR коду
            QRCodeGenerator qrGenerator = new QRCodeGenerator();
            QRCodeData qrCodeData = qrGenerator.CreateQrCode(document.url, QRCodeGenerator.ECCLevel.L);
            PngByteQRCode qrCode = new PngByteQRCode(qrCodeData);
            byte[] qrCodeAsPngByteArr = qrCode.GetGraphic(20);
            #endregion

            #region Вставка QR коду у документ
            var bytes = Convert.FromBase64String(document.base64);

            MemoryStream _mem = new MemoryStream(0);

            _mem.Write(bytes, 0, bytes.Length);

            try
            {
                using (WordprocessingDocument wordDocument = WordprocessingDocument.Open(_mem, true))
                {
                    MainDocumentPart mainPart = wordDocument.MainDocumentPart;

                    if (document.placement == "end")
                    {
                        ImagePart imagePart;
                        imagePart = mainPart.AddImagePart(ImagePartType.Png);
                        using (MemoryStream qrStream = new MemoryStream(qrCodeAsPngByteArr))
                        {
                            imagePart.FeedData(qrStream);
                        }
                        Helpers.AddImageToBody(wordDocument, mainPart.GetIdOfPart(imagePart));

                        // imagePart = GetImagePartByName(wordDocument, "image");

                        //якщо немає QR додємо новий, якщо є оновлюємо старий
                        // if (imagePart == null)
                        // {
                        //   код вище
                        // }
                        // else
                        // {
                        //     using (MemoryStream qrStream = new MemoryStream(qrCodeAsPngByteArr))
                        //     {
                        //         imagePart.FeedData(qrStream);
                        //     }
                        // }

                    }
                    else
                    {
                        mainPart.DeleteParts(mainPart.HeaderParts);

                        HeaderPart header = mainPart.AddNewPart<HeaderPart>();

                        string headerPartId = mainPart.GetIdOfPart(header);

                        ImagePart imagePart = header.AddImagePart(ImagePartType.Png);
                        using (MemoryStream qrStream = new MemoryStream(qrCodeAsPngByteArr))
                        {
                            imagePart.FeedData(qrStream);
                        }

                        Helpers.GenerateHeaderPartContent(header, header.GetIdOfPart(imagePart));
                        IEnumerable<SectionProperties> sections = mainPart.Document.Body.Elements<SectionProperties>();

                        foreach (var section in sections)
                        {
                            section.RemoveAllChildren<HeaderReference>();

                            section.PrependChild<HeaderReference>(new HeaderReference() { Id = headerPartId });
                        }

                    }

                    wordDocument.Save();

                    wordDocument.Close();

                    Result result = new Result("application/vnd.openxmlformats-officedocument.wordprocessingml.document", Convert.ToBase64String(_mem.ToArray()));

                    return result;
                }
                #endregion

                // ImagePart GetImagePartByName(WordprocessingDocument d, string imageName)
                // {
                //     try
                //     {
                //         return d.MainDocumentPart.ImageParts
                //             .Where(p => p.Uri.ToString().Contains(imageName))
                //             .First();
                //     }
                //     catch
                //     {
                //         return null;
                //     }
                // }
            }
            catch (Exception ex)
            {
                return ex;
            }
        }

        [HttpPost("identifier")]
        public ActionResult<object> AddDocumentIdentifier([FromBody] DocumentQR document)
        {

            #region Вставка QR коду у документ
            var bytes = Convert.FromBase64String(document.base64);

            MemoryStream _mem = new MemoryStream(0);

            _mem.Write(bytes, 0, bytes.Length);

            using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(_mem, true))
            {
                var body = wordDoc.MainDocumentPart.Document.Body;

                var paragraphs = body.Elements<DocumentFormat.OpenXml.Wordprocessing.Paragraph>();

                foreach (var para in paragraphs)
                {
                    if (para.InnerText.Contains("ДОГОВІР № ____________"))
                    {
                        foreach (var run in para.Elements<Run>())
                        {
                            foreach (var text in run.Elements<Text>())
                            {
                                if (text.Text.Contains("_"))
                                {
                                    text.Text = new Random().Next(20000, 99999).ToString();
                                }
                            }
                        }
                    }
                }

                wordDoc.Save();

                wordDoc.Close();

                return new Result("application/vnd.openxmlformats-officedocument.wordprocessingml.document", Convert.ToBase64String(_mem.ToArray()));
            }

            #endregion
        }
    }
}

