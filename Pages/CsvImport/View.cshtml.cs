using Microsoft.AspNetCore.Mvc.RazorPages;
using System.IO;
using System.Linq;

namespace NetSuiteAutomation.Pages.CsvImport
{
    public class ViewModel : PageModel
    {
        public string FileName { get; set; }
        public string CsvContent { get; set; }

        public void OnGet()
        {
            string folder = @"C:\ADSK-Automation\NetSuiteImports";

            if (!Directory.Exists(folder))
                return;

            var latestFile = Directory.GetFiles(folder, "*.csv")
                          .OrderByDescending(f => System.IO.File.GetCreationTime(f))
                          .FirstOrDefault();


            if (latestFile == null)
                return;

            FileName = Path.GetFileName(latestFile);
            CsvContent = System.IO.File.ReadAllText(latestFile);
        }
    }
}
