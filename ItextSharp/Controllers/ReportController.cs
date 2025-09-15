using iText.IO.Image;
using iText.Kernel.Colors;
using iText.Kernel.Pdf;
using iText.Kernel.Pdf.Action;
using iText.Kernel.Pdf.Canvas.Draw;
using iText.Layout;
using iText.Layout.Element;
using iText.Layout.Properties;
using Microsoft.AspNetCore.Mvc;
using System.IO;

namespace ItextSharp.Controllers
{
    public class ReportController : Controller
    {
        public IActionResult Header()
        {
            using (var ms = new MemoryStream())
            {
                PdfWriter writer = new PdfWriter(ms);
                PdfDocument pdf = new PdfDocument(writer);
                Document document = new Document(pdf);

                // Header
                Paragraph header = new Paragraph("HEADER")
                    .SetTextAlignment(TextAlignment.CENTER)
                    .SetFontSize(20);

                document.Add(header);

                document.Close();

                // ✅ Return PDF as file download
                return File(ms.ToArray(), "application/pdf", "Header.pdf");
            }
        }

        public IActionResult SubHeader()
        {
            using (var ms = new MemoryStream())
            {
                PdfWriter writer = new PdfWriter(ms);
                PdfDocument pdf = new PdfDocument(writer);
                Document document = new Document(pdf);

                // Header
                Paragraph header = new Paragraph("HEADER")
                    .SetTextAlignment(TextAlignment.CENTER)
                    .SetFontSize(20);

                // Sub Header
                Paragraph subheader = new Paragraph("SUB HEADER")
                   .SetTextAlignment(TextAlignment.CENTER)
                   .SetFontSize(15);

                document.Add(header);
                document.Add(subheader);

                document.Close();

                // ✅ Return PDF as file download
                return File(ms.ToArray(), "application/pdf", "Sub Header.pdf");
            }
        }

        public IActionResult LineSeprator()
        {
            using (var ms = new MemoryStream())
            {
                PdfWriter writer = new PdfWriter(ms);
                PdfDocument pdf = new PdfDocument(writer);
                Document document = new Document(pdf);

                // Header
                Paragraph header = new Paragraph("HEADER")
                    .SetTextAlignment(TextAlignment.CENTER)
                    .SetFontSize(20);

                // Sub Header
                Paragraph subheader = new Paragraph("SUB HEADER")
                   .SetTextAlignment(TextAlignment.CENTER)
                   .SetFontSize(15);

                // Line separator
                LineSeparator ls = new LineSeparator(new SolidLine());

                document.Add(header);
                document.Add(subheader);
                document.Add(ls);

                document.Close();

                // ✅ Return PDF as file download
                return File(ms.ToArray(), "application/pdf", "Line Seperator.pdf");
            }
        }

        public IActionResult AddImage()
        {
            using (var ms = new MemoryStream())
            {
                PdfWriter writer = new PdfWriter(ms);
                PdfDocument pdf = new PdfDocument(writer);
                Document document = new Document(pdf);

                // Header
                Paragraph header = new Paragraph("HEADER")
                    .SetTextAlignment(TextAlignment.CENTER)
                    .SetFontSize(20);

                // Sub Header
                Paragraph subheader = new Paragraph("SUB HEADER")
                   .SetTextAlignment(TextAlignment.CENTER)
                   .SetFontSize(15);

                // Line separator
                LineSeparator ls = new LineSeparator(new SolidLine());

                // Add image
                string imagePath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "img", "logo3.png");

                Image img = new Image(ImageDataFactory
                   .Create(imagePath))
                   .SetTextAlignment(TextAlignment.CENTER);

                img.SetWidth(200);
                img.SetHeight(100);
             // img.ScaleToFit(200, 100);  // fits within 200x100 box keeping aspect ratio
             // img.Scale(0.5f, 0.5f); // 50% of original size

                document.Add(header);
                document.Add(subheader);
                document.Add(ls);
                document.Add(img);

                document.Close();

                // ✅ Return PDF as file download
                return File(ms.ToArray(), "application/pdf", "Add Image.pdf");
            }
        }

        public IActionResult CreatingTable()
        {
            using (var ms = new MemoryStream())
            {
                PdfWriter writer = new PdfWriter(ms);
                PdfDocument pdf = new PdfDocument(writer);
                Document document = new Document(pdf);

                // Header
                Paragraph header = new Paragraph("HEADER")
                    .SetTextAlignment(TextAlignment.CENTER)
                    .SetFontSize(20);

                // Sub Header
                Paragraph subheader = new Paragraph("SUB HEADER")
                   .SetTextAlignment(TextAlignment.CENTER)
                   .SetFontSize(15);

                // Line separator
                LineSeparator ls = new LineSeparator(new SolidLine());

                // Creating Table
                Table table = new Table(2, false);
                Cell cell11 = new Cell(1, 1)
                   .SetBackgroundColor(ColorConstants.GRAY)
                   .SetTextAlignment(TextAlignment.CENTER)
                   .Add(new Paragraph("State"));
                Cell cell12 = new Cell(1, 1)
                   .SetBackgroundColor(ColorConstants.GRAY)
                   .SetTextAlignment(TextAlignment.CENTER)
                   .Add(new Paragraph("Capital"));

                Cell cell21 = new Cell(1, 1)
                   .SetTextAlignment(TextAlignment.CENTER)
                   .Add(new Paragraph("New York"));
                Cell cell22 = new Cell(1, 1)
                   .SetTextAlignment(TextAlignment.CENTER)
                   .Add(new Paragraph("Albany"));

                Cell cell31 = new Cell(1, 1)
                   .SetTextAlignment(TextAlignment.CENTER)
                   .Add(new Paragraph("New Jersey"));
                Cell cell32 = new Cell(1, 1)
                   .SetTextAlignment(TextAlignment.CENTER)
                   .Add(new Paragraph("Trenton"));

                Cell cell41 = new Cell(1, 1)
                   .SetTextAlignment(TextAlignment.CENTER)
                   .Add(new Paragraph("California"));
                Cell cell42 = new Cell(1, 1)
                   .SetTextAlignment(TextAlignment.CENTER)
                   .Add(new Paragraph("Sacramento"));

                table.AddCell(cell11);
                table.AddCell(cell12);
                table.AddCell(cell21);
                table.AddCell(cell22);
                table.AddCell(cell31);
                table.AddCell(cell32);
                table.AddCell(cell41);
                table.AddCell(cell42);

                document.Add(header);
                document.Add(subheader);
                document.Add(ls);
                document.Add(table);

                document.Close();

                // ✅ Return PDF as file download
                return File(ms.ToArray(), "application/pdf", "Creating Table.pdf");
            }
        }

        public IActionResult CreatingHyperlink()
        {
            using (var ms = new MemoryStream())
            {
                PdfWriter writer = new PdfWriter(ms);
                PdfDocument pdf = new PdfDocument(writer);
                Document document = new Document(pdf);

                // Header
                Paragraph header = new Paragraph("HEADER")
                    .SetTextAlignment(TextAlignment.CENTER)
                    .SetFontSize(20);

                // Sub Header
                Paragraph subheader = new Paragraph("SUB HEADER")
                   .SetTextAlignment(TextAlignment.CENTER)
                   .SetFontSize(15);

                // Line separator
                LineSeparator ls = new LineSeparator(new SolidLine());

                // Creating Hyperlink
                Link link = new Link("click here",
               PdfAction.CreateURI("https://www.google.com"));
                link.SetUnderline();
                link.SetFontColor(ColorConstants.BLUE);

                Paragraph hyperLink = new Paragraph("Please ")
                   .Add(link)
                   .Add(" to go www.google.com.");

                document.Add(header);
                document.Add(subheader);
                document.Add(ls);
                document.Add(hyperLink);

                document.Close();

                // ✅ Return PDF as file download
                return File(ms.ToArray(), "application/pdf", "Creating Hyperlink.pdf");
            }
        }

        public IActionResult AddPageNumber()
        {
            using (var ms = new MemoryStream())
            {
                PdfWriter writer = new PdfWriter(ms);
                PdfDocument pdf = new PdfDocument(writer);
                Document document = new Document(pdf);

                // Header
                Paragraph header = new Paragraph("HEADER")
                    .SetTextAlignment(TextAlignment.CENTER)
                    .SetFontSize(20);

                // Sub Header
                Paragraph subheader = new Paragraph("SUB HEADER")
                   .SetTextAlignment(TextAlignment.CENTER)
                   .SetFontSize(15);

                // Line separator
                LineSeparator ls = new LineSeparator(new SolidLine());

                document.Add(header);
                document.Add(subheader);
                document.Add(ls);

                // Add Page Number
                int n = pdf.GetNumberOfPages();
                for (int i = 1; i <= n; i++)
                {
                    document.ShowTextAligned(new Paragraph(String
                       .Format("page" + i + " of " + n)),
                       559, 806, i, TextAlignment.RIGHT,
                       VerticalAlignment.TOP, 0);
                }

                document.Close();

                // ✅ Return PDF as file download
                return File(ms.ToArray(), "application/pdf", "Add Page Number.pdf");
            }
        }
    }
}
