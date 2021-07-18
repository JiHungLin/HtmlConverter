using HtmlToWord.Core;
using HtmlToWord.Service;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Configuration;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Http.Headers;
using System.Text;
using System.Threading.Tasks;

namespace HtmlConverter.Controllers
{
    [ApiController]
    [Route("[controller]")]
    public class ConverterController : Controller
    {
        private const string HtmlWrapper =
            "<!doctype html> <html lang=\"en\"> <head> <meta charset=\"UTF-8\"><title>Document</title> </head><body>{0}</body></html>";
        private readonly IWebHostEnvironment _hostingEnvironment;

        private readonly int DocumentWidth;
        private readonly int DocumentHeight;
        private readonly string _exportFolder;
        private readonly string _inputtFolder;
        private readonly ILogger _logger;
        private readonly IWordApplication _word;
        public ConverterController(IWebHostEnvironment hostingEnvironment, IConfiguration configuration)
        {
            _hostingEnvironment = hostingEnvironment;

            this.DocumentWidth = Convert.ToUInt16(configuration["DocumentWidth"]);
            this.DocumentHeight = Convert.ToUInt16(configuration["DocumentHeight"]);
            this._exportFolder = $@"{hostingEnvironment.ContentRootPath}\Export";
            this._inputtFolder = $@"{hostingEnvironment.ContentRootPath}\Input";
            
            this._logger = new Logger();
            this._word = new WordApplication(this._logger);
            this._word.SetDocumentSize(this.DocumentWidth, this.DocumentHeight);
        }
        public IActionResult Index()
        {
            return View();
        }

        [HttpPost("htmlToWord"), DisableRequestSizeLimit]
        public async Task<IActionResult> HtmlToWord(List<IFormFile> files)
        {
            
            var size = files.Sum(f => f.Length);
            foreach (var file in files)
            {
                if (file.Length > 0)
                {
                    if (file.ContentType != "text/html")
                    {
                        return BadRequest(new { Success = false, Message = "Wrong content type" });
                    }
                    string fileName = file.FileName.Split('.')[0];

                    var inputFilePath = Path.Combine(this._inputtFolder, file.FileName);
                    var exportFilePath = Path.Combine(this._exportFolder, $"{fileName}.doc");
                    var inputFileInfo = new FileInfo(inputFilePath);
                    var exportFileInfo = new FileInfo(exportFilePath);
                    var html = new StringBuilder();
                    using (var reader = new StreamReader(file.OpenReadStream()))
                    {
                        while (reader.Peek() >= 0)
                            html.AppendLine(reader.ReadLine());
                    }
                    try
                    {
                        var htmlFileContent = string.Format(HtmlWrapper, html);
                        System.IO.File.WriteAllText(inputFilePath, htmlFileContent);

                        var success = this._word.ConvertToWord(inputFileInfo, exportFileInfo, out var message);
                        return Ok(new { Success = success, FileUrl = $"{fileName}.doc", Message = message });
                    }
                    catch (Exception e)
                    {
                        this._logger.Info("Failed to export word of {0}", fileName);
                        this._logger.Error("Failed to export word of {0}. Error is {1}", fileName, e);
                        return BadRequest(new { Success = false, Message = e.Message });
                    }
                }
            }

            return BadRequest(new { Success = false, Message = "Please provide one .html file with property name 'files'" });
        }

        [HttpGet("result/{fileName}")]
        public IActionResult Download(string fileName)
        {
            var filePath = $@"{this._exportFolder}\{fileName}";
            return PhysicalFile(filePath, "application/msword");
        }
    }
}
