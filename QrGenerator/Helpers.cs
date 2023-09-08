using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;
using iTextSharp.text;
using iTextSharp.text.pdf;
using Font = iTextSharp.text.Font;
using FontFamily = iTextSharp.text.Font.FontFamily;
using Paragraph = DocumentFormat.OpenXml.Wordprocessing.Paragraph;
using Header = DocumentFormat.OpenXml.Wordprocessing.Header;
using System.Net;
using System;

namespace QrGenerator
{
    public class Helpers
    {
        public static void AddWaterMark(string watermarkText, PdfReader reader, PdfStamper stamper)
        {
            int pages = reader.NumberOfPages;
            Font f = new Font(FontFamily.HELVETICA, 72);
            Phrase p = new Phrase(watermarkText, f);

            PdfGState gs1 = new PdfGState();
            gs1.FillOpacity = 0.1f;

            PdfContentByte over;
            Rectangle pagesize;

            float x, y;
            for (int i = 1; i <= pages; i++)
            {
                pagesize = reader.GetPageSize(i);
                x = (pagesize.Left + pagesize.Right) / 2;
                y = (pagesize.Top + pagesize.Bottom) / 2;
                over = stamper.GetOverContent(i);
                over.SaveState();
                over.SetGState(gs1);

                ColumnText.ShowTextAligned(over, Element.ALIGN_CENTER, p, x, y, 20);

                over.RestoreState();
            }
        }

        public static string TinyLink(string link)
        {
            string tinyURL;

            try
            {
                Uri tinyUrlAPI = new Uri("http://tinyurl.com/api-create.php?url=" + link);
                WebClient client = new WebClient();
                tinyURL = client.DownloadString(tinyUrlAPI);
            }
            catch
            {
                tinyURL = link;
            }

            return tinyURL;
        }
    }
}
