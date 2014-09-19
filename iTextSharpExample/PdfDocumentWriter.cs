using System.Collections.Generic;
using System.Drawing;
using System.IO;
using iTextSharp.text;
using iTextSharp.text.pdf;
using Font = iTextSharp.text.Font;

namespace iTextSharpExample
{
    public interface IPdfDocumentWriter
    {
        MemoryStream CreateCompanyRegistrationDocument(string fileName, string companyName, string companyNumber);
    }

    public class PdfDocumentWriter : IPdfDocumentWriter
    {
        private const int textSize = 8;
        private const float textRotation = 0;
        private const string userPassword = "";
        private const string ownerPassword = "secret password";
        private const int permision = PdfWriter.AllowPrinting;

        readonly Dictionary<int, Point> companyNameLocations = new Dictionary<int, Point>
            {
                { 1, new Point(140,740) },
                { 3, new Point(140,310) }
            };

        readonly Dictionary<int, Point> companyNumberLocations = new Dictionary<int, Point>
            {
                { 1, new Point(460,740) },
                { 3, new Point(460,310) }
            };


        public MemoryStream CreateCompanyRegistrationDocument(string fileName, string companyName, string companyNumber)
        {
            using (var stumpedDocStream = new MemoryStream())
            {
                PdfReader reader = null;
                PdfStamper stamper = null;
                try
                {
                    reader = new PdfReader(fileName);
                    stamper = new PdfStamper(reader, stumpedDocStream);
                    var bf = BaseFont.CreateFont(BaseFont.HELVETICA, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
                    var font = new Font(bf, textSize);
                    for (var pageNumber = 1; pageNumber < reader.NumberOfPages + 1; pageNumber++)
                    {
                        var canvas = stamper.GetOverContent(pageNumber);
                        canvas.SetFontAndSize(bf, textSize);
                        RenderPhase(pageNumber, companyNameLocations, canvas, companyName, font);
                        RenderPhase(pageNumber, companyNumberLocations, canvas, companyNumber, font);
                    }
                }
                finally
                {
                    try
                    {
                        if(reader!=null)
                            reader.Close();
                    }
                    finally
                    {
                        if (stamper != null)
                            stamper.Close();
                    }
                }

                return new MemoryStream(EncryptPdf(stumpedDocStream.ToArray()));
            }
        }

        private static void RenderPhase(int pageNumber, Dictionary<int, Point> locations, PdfContentByte canvas, string phase, Font font)
        {
            Point position;
            locations.TryGetValue(pageNumber, out position);
            if (!position.IsEmpty)
            {
                ColumnText.ShowTextAligned(canvas, Element.ALIGN_LEFT, new Phrase(phase, font), position.X, position.Y, textRotation);
            }
        }

        private static byte[] EncryptPdf(byte[] pdfDoc)
        {
            PdfReader reader = null;
            try
            {
                reader = new PdfReader(pdfDoc);
                using (var output = new MemoryStream())
                {
                    PdfEncryptor.Encrypt(reader, output, true, userPassword, ownerPassword, permision);
                    return output.ToArray();
                }
            }
            finally
            {
                if (reader != null)
                    reader.Close();
            }
        }
    }
}
