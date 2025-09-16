using iText.IO.Image;
using iText.Kernel.Colors;
using iText.Kernel.Font;
using iText.Kernel.Geom;
using iText.Kernel.Pdf;
using iText.Kernel.Pdf.Action;
using iText.Kernel.Pdf.Canvas.Draw;
using iText.Layout;
using iText.Layout.Borders;
using iText.Layout.Element;
using iText.Layout.Properties;
using iText.StyledXmlParser.Jsoup.Nodes;
using Microsoft.AspNetCore.Mvc;
using System.IO;
using Document = iText.Layout.Document;

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
                string imagePath = System.IO.Path.Combine(
                 Directory.GetCurrentDirectory(), "wwwroot", "img", "logo3.png");


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

        public IActionResult DailySalesReport()
        {
            using (var ms = new MemoryStream())
            {
                PdfWriter writer = new PdfWriter(ms);
                PdfDocument pdf = new PdfDocument(writer);
                Document document = new Document(pdf, iText.Kernel.Geom.PageSize.A4);

                Paragraph header = new Paragraph("🍴 Umesh Restaurant")
                    .SetTextAlignment(TextAlignment.CENTER)
                    .SetFontSize(20);
                // .SetBold() removed -- not available for Paragraph

                Paragraph subHeader = new Paragraph("Daily Sales Report - 15 Sept 2025")
                    .SetTextAlignment(TextAlignment.CENTER)
                    .SetFontSize(14)
                    .SetFontColor(ColorConstants.DARK_GRAY);

                document.Add(header);
                document.Add(subHeader);
                document.Add(new LineSeparator(new SolidLine()));

                // ====== TABLE ======
                Table table = new Table(new float[] { 3, 1, 2, 2 }).UseAllAvailableWidth();
                table.AddHeaderCell(new Cell().Add(new Paragraph("Item")).SetBackgroundColor(ColorConstants.LIGHT_GRAY));
                table.AddHeaderCell(new Cell().Add(new Paragraph("Qty")).SetBackgroundColor(ColorConstants.LIGHT_GRAY));
                table.AddHeaderCell(new Cell().Add(new Paragraph("Price")).SetBackgroundColor(ColorConstants.LIGHT_GRAY));
                table.AddHeaderCell(new Cell().Add(new Paragraph("Total")).SetBackgroundColor(ColorConstants.LIGHT_GRAY));

                AddRow(table, "Paneer Butter Masala", 5, 250);
                AddRow(table, "Dal Tadka", 8, 150);
                AddRow(table, "Butter Naan", 20, 40);
                AddRow(table, "Cold Drink", 12, 50);

                table.AddCell(new Cell(1, 3).Add(new Paragraph("Grand Total")));
                table.AddCell(new Cell().Add(new Paragraph("₹ 3,870")));

                document.Add(table);

                document.Add(new Paragraph("\nReport Generated On: " + DateTime.Now.ToString("dd-MMM-yyyy HH:mm"))
                    .SetFontSize(10)
                    .SetFontColor(ColorConstants.GRAY));

                // ====== FOOTER (Page Numbers) ======
                int n = pdf.GetNumberOfPages();
                for (int i = 1; i <= n; i++)
                {
                    document.ShowTextAligned(
                        new Paragraph($"Page {i} of {n}"),
                        559, 20, i, // bottom-right corner
                        TextAlignment.RIGHT,
                        VerticalAlignment.BOTTOM,
                        0
                    );
                }

                document.Close();

                return File(ms.ToArray(), "application/pdf", "DailySalesReport.pdf");
            }
        }

        // Helper function for table rows
        private void AddRow(Table table, string item, int qty, decimal price)
        {
            decimal total = qty * price;
            table.AddCell(new Cell().Add(new Paragraph(item)));
            table.AddCell(new Cell().Add(new Paragraph(qty.ToString())));
            table.AddCell(new Cell().Add(new Paragraph("₹ " + price)));
            table.AddCell(new Cell().Add(new Paragraph("₹ " + total)));
        }

        public IActionResult BillReport()
        {
            using var ms = new MemoryStream();
            PdfWriter writer = new PdfWriter(ms);
            PdfDocument pdf = new PdfDocument(writer);
            Document doc = new Document(pdf);

            // Header
            doc.Add(new Paragraph("Restaurant Bill")
                .SetTextAlignment(TextAlignment.CENTER)
                .SetFontSize(20)
                );

            doc.Add(new Paragraph("Date: 15-Sep-2025").SetFontSize(12));

            // Table
            Table table = new Table(3, true);
            table.AddHeaderCell("Item");
            table.AddHeaderCell("Qty");
            table.AddHeaderCell("Price");

            table.AddCell("Burger");
            table.AddCell("2");
            table.AddCell("$10");

            table.AddCell("Pizza");
            table.AddCell("1");
            table.AddCell("$15");

            table.AddCell("Coke");
            table.AddCell("3");
            table.AddCell("$6");

            doc.Add(table);
            doc.Add(new Paragraph("Total: $31").SetTextAlignment(TextAlignment.RIGHT));

            doc.Close();
            return File(ms.ToArray(), "application/pdf", "BillReport.pdf");
        }

        public IActionResult SalarySlip()
        {
            using var ms = new MemoryStream();
            PdfWriter writer = new PdfWriter(ms);
            PdfDocument pdf = new PdfDocument(writer);
            Document doc = new Document(pdf);

            doc.Add(new Paragraph("Salary Slip")
                .SetTextAlignment(TextAlignment.CENTER)
                .SetFontSize(20)
                );

            doc.Add(new Paragraph("Employee: Umesh Kumar\nMonth: August 2025").SetFontSize(12));

            Table table = new Table(2, true);
            table.AddCell("Basic Salary");
            table.AddCell("$2000");
            table.AddCell("HRA");
            table.AddCell("$500");
            table.AddCell("Allowances");
            table.AddCell("$300");
            table.AddCell("Deductions");
            table.AddCell("($200)");
            table.AddCell("Net Salary");
            table.AddCell("$2600");

            doc.Add(table);
            doc.Close();
            return File(ms.ToArray(), "application/pdf", "SalarySlip.pdf");
        }

        public IActionResult Invoice()
        {
            using var ms = new MemoryStream();
            PdfWriter writer = new PdfWriter(ms);
            PdfDocument pdf = new PdfDocument(writer);
            Document doc = new Document(pdf);

            doc.Add(new Paragraph("Invoice")
                .SetTextAlignment(TextAlignment.CENTER)
                .SetFontSize(20)
                );

            doc.Add(new Paragraph("Invoice No: INV-1001\nDate: 15-Sep-2025").SetFontSize(12));

            Table table = new Table(4, true);
            table.AddHeaderCell("Product");
            table.AddHeaderCell("Qty");
            table.AddHeaderCell("Unit Price");
            table.AddHeaderCell("Total");

            table.AddCell("Laptop");
            table.AddCell("1");
            table.AddCell("$800");
            table.AddCell("$800");

            table.AddCell("Mouse");
            table.AddCell("2");
            table.AddCell("$20");
            table.AddCell("$40");

            table.AddCell("Keyboard");
            table.AddCell("1");
            table.AddCell("$40");
            table.AddCell("$40");

            doc.Add(table);
            doc.Add(new Paragraph("Grand Total: $880").SetTextAlignment(TextAlignment.RIGHT));

            doc.Close();
            return File(ms.ToArray(), "application/pdf", "Invoice.pdf");
        }

        public IActionResult certificate()
        {
            using var ms = new MemoryStream();
            PdfWriter writer = new PdfWriter(ms);
            PdfDocument pdf = new PdfDocument(writer);
            Document doc = new Document(pdf);

            doc.Add(new Paragraph("Certificate of Achievement")
                .SetTextAlignment(TextAlignment.CENTER)
                .SetFontSize(24)
              );

            doc.Add(new Paragraph("\n\nThis is to certify that")
                .SetTextAlignment(TextAlignment.CENTER)
                .SetFontSize(14));

            doc.Add(new Paragraph("Umesh Kumar Singh")
                .SetTextAlignment(TextAlignment.CENTER)
                .SetFontSize(20)
                );

            doc.Add(new Paragraph("\nHas successfully completed the training in .NET & Angular")
                .SetTextAlignment(TextAlignment.CENTER)
                .SetFontSize(14));

            doc.Add(new Paragraph("\nDate: 15-Sep-2025")
                .SetTextAlignment(TextAlignment.CENTER)
                .SetFontSize(12));

            doc.Add(new Paragraph("\n\n____________________          ____________________")
                .SetTextAlignment(TextAlignment.CENTER)
                .SetFontSize(12));

            doc.Add(new Paragraph("Authorized Signatory                     HR Manager")
                .SetTextAlignment(TextAlignment.CENTER)
                .SetFontSize(12));

            doc.Close();
            return File(ms.ToArray(), "application/pdf", "Certificate.pdf");
        }

        public IActionResult Sample1()
        {
            using var ms = new MemoryStream();
            PdfWriter writer = new PdfWriter(ms);
            PdfDocument pdf = new PdfDocument(writer);
            Document doc = new Document(pdf, PageSize.A3);
            doc.SetMargins(20f, 20f, 60f, 40f);

            // ====== HEADER ======
            Paragraph header = new Paragraph("Aimil Calibration Laboratory\nNew Delhi")
                .SetTextAlignment(TextAlignment.CENTER)
                //.SetBold()
                .SetFontSize(14);
            doc.Add(header);

            Paragraph subHeader = new Paragraph("Master Sample Register")
                .SetTextAlignment(TextAlignment.CENTER)
                //.SetBold()
                .SetFontSize(12);
            doc.Add(subHeader);

            Paragraph meta = new Paragraph(
                "MSP No: ACL/MSP/08   Issue No: 02   Issue Date: 26.07.2022   " +
                "Amend No: 00   Amend Date: ---\nStandard ISO/IEC 17025:2017 Clause No: 7.4.2   Format No: ACL-MSP/08/F-01"
            ).SetTextAlignment(TextAlignment.CENTER).SetFontSize(9);
            doc.Add(meta);

            doc.Add(new Paragraph("\n"));

            // ====== TABLE ======
            Table table = new Table(new float[] { 35, 90, 150, 105, 160, 30, 70, 90, 110, 90, 110 })
                .UseAllAvailableWidth();

            string[] headers = {
        "S.No", "Date of Receipt", "Customer", "Request Letter Ref", "Name of Instrument",
        "Qty", "Calibration Method", "Job no. / Lab Code", "Certificate No.",
        "Date of Return", "Receiver’s Name"
    };

            foreach (var h in headers)
            {
                table.AddHeaderCell(new Cell().Add(new Paragraph(h))
                    .SetBackgroundColor(ColorConstants.LIGHT_GRAY)
                    .SetTextAlignment(TextAlignment.CENTER)
                    .SetFontSize(9));
                    //.SetBold());
            }

            // Static sample rows
            table.AddCell("1");
            table.AddCell("02-09-2025");
            table.AddCell("Anuj Sharma");
            table.AddCell("Aimil/007 (On Site)");
            table.AddCell("LENGTH GAUGE (451)");
            table.AddCell("1");
            table.AddCell("DM108");
            table.AddCell("ACL/2025-26/9/DM108");
            table.AddCell("");
            table.AddCell("");
            table.AddCell("");

            table.AddCell("2");
            table.AddCell("02-09-2025");
            table.AddCell("Anuj Sharma");
            table.AddCell("Aimil/007 (On Site)");
            table.AddCell("LENGTH GAUGE (451)");
            table.AddCell("1");
            table.AddCell("DM109");
            table.AddCell("ACL/2025-26/9/DM109");
            table.AddCell("");
            table.AddCell("");
            table.AddCell("");

            doc.Add(table);

            doc.Add(new Paragraph("\nCONTROLLED COPY")
                .SetTextAlignment(TextAlignment.CENTER)
                .SetFontSize(10));
                //.SetBold());

            // ====== FOOTER SIGNATURES ======
            Table footer = new Table(3).UseAllAvailableWidth();
            footer.AddCell(new Cell().Add(new Paragraph("Prepared By\nSystem Manager - ACL"))
                .SetTextAlignment(TextAlignment.CENTER).SetBorder(Border.NO_BORDER));
            footer.AddCell(new Cell().Add(new Paragraph("Reviewed & Approved By\nHead - ACL"))
                .SetTextAlignment(TextAlignment.CENTER).SetBorder(Border.NO_BORDER));
            footer.AddCell(new Cell().Add(new Paragraph("Issued By\nTechnical Manager - ACL"))
                .SetTextAlignment(TextAlignment.CENTER).SetBorder(Border.NO_BORDER));

            doc.Add(new Paragraph("\n\n"));
            doc.Add(footer);

            doc.Close();
            return File(ms.ToArray(), "application/pdf", "StaticMasterSampleRegister.pdf");
        }
    }
}
