using System.Net;
using Microsoft.AspNetCore.Mvc;

namespace DocsConverter.Controllers
{
  [ApiController]
  [Route("[controller]")]
  public class ConvertController(Hwp2Pdf _hwp2Pdf) : ControllerBase
  {
    private readonly Hwp2Pdf hwp2Pdf = _hwp2Pdf;
    readonly string basePath = AppDomain.CurrentDomain.BaseDirectory + @"\";

    [HttpGet]
    public IActionResult Get1([FromQuery] string? fileUrl)
    {
      hwp2Pdf.Convert_file("input.hwp", "output.pdf");
      if (string.IsNullOrEmpty(fileUrl))
      {
        return Ok("Please provide a file URL.");
      }
      return Ok($"Converting file at {fileUrl}...");
    }

    [HttpGet("default")]
    public IActionResult GetDefault()
    {
      return Ok("Hello, World!2");
    }

    private void DownloadAndConvertFile(string fileUrl, string outputFilename, string outputMimeType, out string filePath, out string fileName, out long fileSize)
    {
      string basePath = this.basePath + @"files\";
      string inputFilename = "input" + Path.GetExtension(new Uri(fileUrl).LocalPath);
      using (HttpClient client = new())
      {
        var response = client.GetAsync(fileUrl).Result;
        response.EnsureSuccessStatusCode();
        var fileBytes = response.Content.ReadAsByteArrayAsync().Result;
        System.IO.File.WriteAllBytes(basePath + inputFilename, fileBytes);
      }
      Console.WriteLine("Downloaded file to " + basePath + inputFilename);
      hwp2Pdf.Convert_file(inputFilename, outputFilename);
      filePath = basePath + outputFilename;
      fileName = Path.GetFileNameWithoutExtension(new Uri(fileUrl).LocalPath) + Path.GetExtension(outputFilename);
      fileSize = new FileInfo(filePath).Length;
    }

    [HttpGet("to-thumbnail")]
    public IActionResult GetThumbnail([FromQuery] string? fileUrl)
    {
      if (string.IsNullOrEmpty(fileUrl))
      {
        return BadRequest("Please provide a file URL.");
      }
      string basePath = this.basePath + @"files\";
      string inputFilename = "input" + Path.GetExtension(new Uri(fileUrl).LocalPath);
      using (HttpClient client = new())
      {
        var response = client.GetAsync(fileUrl).Result;
        response.EnsureSuccessStatusCode();
        var fileBytes = response.Content.ReadAsByteArrayAsync().Result;
        System.IO.File.WriteAllBytes(basePath + inputFilename, fileBytes);
      }
      hwp2Pdf.Convert_file(inputFilename, "output.png");
      string documentsPath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
      string filePath = Path.Combine(documentsPath, "image001.png");
      string fileName = Path.GetFileNameWithoutExtension(new Uri(fileUrl).LocalPath) + Path.GetExtension(filePath);
      long fileSize = new FileInfo(filePath).Length;

      var fileStream = new FileStream(filePath, FileMode.Open, FileAccess.Read);
      Response.Headers.Append("Content-Disposition", $"filename={fileName}");
      Response.Headers.Append("Content-Length", fileSize.ToString());
      return new FileStreamResult(fileStream, "image/png");
    }

    [HttpGet("to-pdf")]
    public IActionResult GetPdf([FromQuery] string? fileUrl)
    {
      if (string.IsNullOrEmpty(fileUrl))
      {
        return BadRequest("Please provide a file URL.");
      }

      DownloadAndConvertFile(fileUrl, "output.pdf", "application/pdf", out var filePath, out var fileName, out var fileSize);
      var fileStream = new FileStream(filePath, FileMode.Open, FileAccess.Read);
      Response.Headers.Append("Content-Disposition", $"filename={fileName}");
      Response.Headers.Append("Content-Length", fileSize.ToString());
      return new FileStreamResult(fileStream, "application/pdf");
    }

    [HttpGet("to-docx")]
    public IActionResult GetDocx([FromQuery] string? fileUrl)
    {
      if (string.IsNullOrEmpty(fileUrl))
      {
        return BadRequest("Please provide a file URL.");
      }

      DownloadAndConvertFile(fileUrl, "output.docx", "application/vnd.openxmlformats-officedocument.wordprocessingml.document", out var filePath, out var fileName, out var fileSize);
      var fileStream = new FileStream(filePath, FileMode.Open, FileAccess.Read);
      Response.Headers.Append("Content-Disposition", $"filename={fileName}");
      Response.Headers.Append("Content-Length", fileSize.ToString());
      return new FileStreamResult(fileStream, "application/vnd.openxmlformats-officedocument.wordprocessingml.document");
    }

    [HttpGet("to-html")]
    public IActionResult GetHtml([FromQuery] string? fileUrl)
    {
      if (string.IsNullOrEmpty(fileUrl))
      {
        return BadRequest("Please provide a file URL.");
      }

      DownloadAndConvertFile(fileUrl, "output.html", "text/html", out var filePath, out var fileName, out var fileSize);
      var fileStream = new FileStream(filePath, FileMode.Open, FileAccess.Read);
      Response.Headers.Append("Content-Disposition", $"filename={fileName}");
      Response.Headers.Append("Content-Length", fileSize.ToString());
      return new FileStreamResult(fileStream, "text/html");
    }

    [HttpGet("to-txt")]
    public IActionResult GetTxt([FromQuery] string? fileUrl)
    {
      if (string.IsNullOrEmpty(fileUrl))
      {
        return BadRequest("Please provide a file URL.");
      }

      DownloadAndConvertFile(fileUrl, "output.txt", "text/plain", out var filePath, out var fileName, out var fileSize);
      var fileStream = new FileStream(filePath, FileMode.Open, FileAccess.Read);
      Response.Headers.Append("Content-Disposition", $"filename={fileName}");
      Response.Headers.Append("Content-Length", fileSize.ToString());
      return new FileStreamResult(fileStream, "text/plain");
    }

    [HttpHead("to-txt")]
    public IActionResult HeadTxt([FromQuery] string? fileUrl)
    {
      if (string.IsNullOrEmpty(fileUrl))
      {
        return BadRequest("Please provide a file URL.");
      }

      DownloadAndConvertFile(fileUrl, "output.txt", "text/plain", out var filePath, out var fileName, out var fileSize);
      var fileStream = new FileStream(filePath, FileMode.Open, FileAccess.Read);
      Response.Headers.Append("Content-Disposition", $"filename={fileName}");
      Response.Headers.Append("Content-Length", fileSize.ToString());
      return new FileStreamResult(fileStream, "text/plain");
    }

    [HttpGet("to-rtf")]
    public IActionResult GetRtf([FromQuery] string? fileUrl)
    {
      if (string.IsNullOrEmpty(fileUrl))
      {
        return BadRequest("Please provide a file URL.");
      }

      DownloadAndConvertFile(fileUrl, "output.rtf", "application/rtf", out var filePath, out var fileName, out var fileSize);
      var fileStream = new FileStream(filePath, FileMode.Open, FileAccess.Read);
      Response.Headers.Append("Content-Disposition", $"attachment; filename={fileName}");
      Response.Headers.Append("Content-Length", fileSize.ToString());
      return new FileStreamResult(fileStream, "application/rtf");
    }

    [HttpGet("to-hwp")]
    public IActionResult GetHwp([FromQuery] string? fileUrl)
    {
      if (string.IsNullOrEmpty(fileUrl))
      {
        return BadRequest("Please provide a file URL.");
      }

      DownloadAndConvertFile(fileUrl, "output.hwp", "application/octet-stream", out var filePath, out var fileName, out var fileSize);
      var fileStream = new FileStream(filePath, FileMode.Open, FileAccess.Read);
      Response.Headers.Append("Content-Disposition", $"attachment; filename={fileName}");
      Response.Headers.Append("Content-Length", fileSize.ToString());
      return new FileStreamResult(fileStream, "application/octet-stream");
    }
  }
}