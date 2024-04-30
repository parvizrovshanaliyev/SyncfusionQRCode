using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.DocIORenderer;
using Syncfusion.Drawing;
using Syncfusion.Pdf;
using Syncfusion.Pdf.Barcode;
using Syncfusion.Pdf.Parsing;

namespace SyncfusionQRCode.Services;

public interface ISyncfusionService
{
    byte[] AddQrCodeToPdf();
    byte[] AddQrCodeToWord();

    byte[] AddQrCodeToPdfDocument();
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
        using (FileStream docStream = new FileStream(Path.GetFullPath(pdfFilePath), FileMode.Open, FileAccess.Read))
        {
            //Loads file stream into Word document
            using (WordDocument wordDocument = new WordDocument(docStream, FormatType.Automatic))
            {
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
        }
    }
}