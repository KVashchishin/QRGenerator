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
using System.Net;

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

                        Helpers.AddWaterMark(pdf.watermark, reader, stamper);
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

                        Helpers.AddWaterMark(pdf.watermark, reader, stamper);
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

        [HttpGet("test")]
        public ActionResult<object> TestGet()
        {
            try
            {
                WebRequest request = WebRequest.Create(
                  "https://docs.microsoft.com");
                request.Credentials = CredentialCache.DefaultCredentials;

                WebResponse response = request.GetResponse();

                string status;
                string responseFromServer;

                using (Stream dataStream = response.GetResponseStream())
                {
                    StreamReader reader = new StreamReader(dataStream);
                    status = ((HttpWebResponse)response).StatusDescription;
                    responseFromServer = reader.ReadToEnd();
                }

                // Close the response.
                response.Close();

                return new Result(status, responseFromServer);
            }
            catch (Exception ex)
            {
                return ex;
            }
        }
    }
}

