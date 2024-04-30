using Microsoft.AspNetCore.Mvc;
using SyncfusionQRCode.Services;

namespace SyncfusionQRCode.Controllers;

[ApiController]
[Route("[controller]")]
public class QRCodeController : ControllerBase
{
    private readonly ILogger<QRCodeController> _logger;
    private readonly ISyncfusionService _syncfusionService;

    public QRCodeController(ILogger<QRCodeController> logger, ISyncfusionService syncfusionService)
    {
        _logger = logger;
        _syncfusionService = syncfusionService;
    }

    [HttpGet("AddQRCodeToPDF")]
    public IActionResult AddQRCodeToPDF()
    {
        var byteArray = _syncfusionService.AddQrCodeToPdf();
        
        return File(byteArray, "application/pdf", "QRCode.pdf");
    }
    
    [HttpGet("AddQrCodeToPdfDocument")]
    public IActionResult AddQrCodeToPdfDocument()
    {
        var byteArray = _syncfusionService.AddQrCodeToPdfDocument();
        
        return File(byteArray, "application/pdf", "QRCodeDocument.pdf");
    }
    
    // [HttpGet(Name = "AddQRCodeToWord")]
    // public IActionResult AddQRCodeToWord()
    // {
    //     var byteArray = _syncfusionService.AddQrCodeToWord();
    //     
    //     return File(byteArray, "application/vnd.openxmlformats-officedocument.wordprocessingml.document", "QRCode.docx");
    // }
        
}