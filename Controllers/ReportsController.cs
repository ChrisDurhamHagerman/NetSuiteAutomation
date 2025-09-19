using Microsoft.AspNetCore.Mvc;
using System.IO;
using System.IO.Compression;

namespace NetSuiteAutomation.Controllers
{
    [ApiController]
    [Route("api/[controller]")]
    public class ReportsController : ControllerBase
    {
        private const string ReportsFolder = @"C:\ADSK-Automation\EmailReports";

        [HttpGet("download")]
        public IActionResult Download()
        {
            // 1) Grab whatever .xlsx files are in EmailReports
            if (!Directory.Exists(ReportsFolder))
                return NotFound("Reports folder not found.");

            var files = Directory.GetFiles(ReportsFolder, "*.xlsx");
            if (files.Length == 0)
                return NotFound("No report files available.");

            // 2) Zip them up
            using var ms = new MemoryStream();
            using (var zip = new ZipArchive(ms, ZipArchiveMode.Create, true))
            {
                foreach (var f in files)
                {
                    var entry = zip.CreateEntry(Path.GetFileName(f));
                    using var fs = System.IO.File.OpenRead(f);
                    using var es = entry.Open();
                    fs.CopyTo(es);
                }
            }
            ms.Position = 0;

            // 3) Return ZIP
            return File(ms.ToArray(),
                        "application/zip",
                        "EmailReports.zip");
        }
    }
}
