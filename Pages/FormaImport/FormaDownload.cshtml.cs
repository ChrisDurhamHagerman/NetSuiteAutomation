using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;

namespace NetSuiteAutomation.Pages
{
    public class FormaDownloadModel : PageModel
    {
        private readonly string _downloadPath = @"C:\ADSK-Automation\FormaDownloads\all_transaction_fees.csv";

        [BindProperty]
        public string StatusMessage { get; set; }

        public IActionResult OnGetDownload()
        {
            if (!System.IO.File.Exists(_downloadPath))
            {
                StatusMessage = "❌ No Forma CSV found at the expected path.";
                return Page();
            }

            byte[] fileBytes = System.IO.File.ReadAllBytes(_downloadPath);
            return File(fileBytes, "text/csv", "all_transaction_fees.csv");
        }

        public void OnGet()
        {
            StatusMessage = "";
        }
    }
}
