using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.DocIORenderer;
using Syncfusion.Drawing;
using Syncfusion.Pdf;
using Syncfusion.Pdf.Barcode;
using Syncfusion.Pdf.Graphics;
using Syncfusion.Pdf.Parsing;

namespace SyncfusionQRCode.Services;

public interface ISyncfusionService
{
    byte[] AddQrCodeToPdf();
    byte[] AddQrCodeToWord();

    byte[] AddQrCodeToPdfDocument();

    byte[] AddQrCodeToPdfFooter();
}

public class SyncfusionService : ISyncfusionService
{
    private readonly IWebHostEnvironment _hostingEnvironment;

    public SyncfusionService(IWebHostEnvironment hostingEnvironment)
    {
        _hostingEnvironment = hostingEnvironment;
    }

    public byte[] AddQrCodeToPdf()
    {
        // Create a new memory stream
        using var stream = new MemoryStream();

        //Create a new PDF document
        PdfDocument document = new PdfDocument();

        //Add a page to the document
        PdfPage page = document.Pages.Add();

        //Create a QR code
        PdfQRBarcode qrCode = new PdfQRBarcode();

        //Set the QR code text
        qrCode.Text = "Test QR Code";

        //Set the QR code size
        qrCode.XDimension = 3;

        //Draw the QR code on the page
        qrCode.Draw(page, new PointF(10, 10));

        //Save the document to the memory stream
        document.Save(stream);

        //Close the document
        document.Close(true);

        // Convert the memory stream to a byte array
        byte[] byteArray = stream.ToArray();

        return byteArray;
    }

    public byte[] AddQrCodeToWord()
    {
        throw new NotImplementedException();
    }

    public byte[] AddQrCodeToPdfDocument()
    {
        var pdfFilePath = Path.Combine(_hostingEnvironment.WebRootPath, "QRCode.docx");

        //Open the file as Stream
        using FileStream docStream = new FileStream(Path.GetFullPath(pdfFilePath), FileMode.Open, FileAccess.Read);

        //Loads file stream into Word document
        using WordDocument wordDocument = new WordDocument(docStream, FormatType.Automatic);
        //Instantiation of DocIORenderer for Word to PDF conversion
        using (DocIORenderer render = new DocIORenderer())
        {
            //Converts Word document into PDF document
            PdfDocument pdfDocument = render.ConvertToPDF(wordDocument);

            //Saves the PDF document to MemoryStream.
            MemoryStream stream = new MemoryStream();
            pdfDocument.Save(stream);
            stream.Position = 0;

            //Create Qr Code
            PdfQRBarcode qrCode = new PdfQRBarcode();

            //Set the QR code text
            qrCode.Text = "Test QR Code";

            //Set the QR code size
            qrCode.XDimension = 3;

            //Draw the QR code on the page footer
            qrCode.Draw(pdfDocument.Pages[0], new PointF(100, 100), new SizeF(100, 100));

            //Save the document to the memory stream
            pdfDocument.Save(stream);

            //Close the document
            pdfDocument.Close(true);

            // Convert the memory stream to a byte array
            byte[] byteArray = stream.ToArray();

            // this is the byte array to word document
            return byteArray;
        }
    }


    public byte[] AddQrCodeToPdfFooter()
    {
        var templatePath = Path.Combine(_hostingEnvironment.WebRootPath, "pdfTemplate.pdf");

        // Open the file as Stream
        using (FileStream pdfFileStream =
               new FileStream(Path.GetFullPath(templatePath), FileMode.Open, FileAccess.Read, FileShare.Read))
        {
            // Load the existing PDF document
            PdfLoadedDocument loadedDocument = new PdfLoadedDocument(pdfFileStream);

            // Create a QR code
            PdfQRBarcode qrCode = new PdfQRBarcode();

            // Set the QR code text
            qrCode.Text = "Test QR Code";

            // Set the QR code size
            qrCode.XDimension = 3;

            // Calculate the position for the QR code in the footer
            PointF qrCodePosition = new PointF(loadedDocument.Pages[0].Size.Width / 2 - qrCode.XDimension * 5, 20);

            // Draw the QR code on the page
            qrCode.Draw(loadedDocument.Pages[0], qrCodePosition);

            // Save the document to the memory stream
            MemoryStream stream = new MemoryStream();
            loadedDocument.Save(stream);

            // Close the document
            loadedDocument.Close(true);

            // Convert the memory stream to a byte array
            byte[] byteArray = stream.ToArray();

            return byteArray;
        }
    }


    // PDF-Examples/Header and Footer/Adding-an-automatic-field-in-header-and-footer
    public byte[] AddQrCodeToPdfFooter2()
    {
        var templatePath = Path.Combine(_hostingEnvironment.WebRootPath, "pdfTemplate.pdf");

        // Open the file as Stream
        using FileStream pdfFileStream =
            new FileStream(Path.GetFullPath(templatePath), FileMode.Open, FileAccess.Read, FileShare.Read);

        //Load the PDF document. 
        PdfLoadedDocument loadedDocument = new PdfLoadedDocument(pdfFileStream);

        //Load the page.
        PdfLoadedPage loadedPage = loadedDocument.Pages[0] as PdfLoadedPage;

        //Create the template from the page.
        PdfTemplate template = loadedPage.CreateTemplate();

        //Create a new PDF document.
        PdfDocument document = new PdfDocument();

        //Set the document margin.
        document.PageSettings.SetMargins(2);

        //Add the page.
        PdfPage page = document.Pages.Add();

        //Create the graphics.
        PdfGraphics graphics = page.Graphics;

        //Draw the template.
        graphics.DrawPdfTemplate(template, Syncfusion.Drawing.PointF.Empty, new Syncfusion.Drawing.SizeF(page.Size.Width / 2, page.Size.Height));
        
        // create a QR code
        PdfQRBarcode qrCode = new PdfQRBarcode();
        
        // set the QR code text
        qrCode.Text = "Test QR Code";
        
        // set the QR code size
        qrCode.XDimension = 3;
        
        // // calculate the position for the QR code in the footer
        // PointF qrCodePosition = new PointF(loadedDocument.Pages[0].Size.Width / 2 - qrCode.XDimension * 5, 20);
        //
        // // draw the QR code on the page
        // qrCode.Draw(page, qrCodePosition);
        
        // document.Template.
         
        
        // return the byte array
        // Save the document to the memory stream
        MemoryStream stream = new MemoryStream();
        
        document.Save(stream);
        
        //Close the documents.
        loadedDocument.Close(true);
        document.Close(true);
        
        // Convert the memory stream to a byte array
        byte[] byteArray = stream.ToArray();

        return byteArray;
    }
}