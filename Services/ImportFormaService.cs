using System;
using System.IO;
using System.Net.Http;
using System.Threading.Tasks;
using NetSuiteAutomation.Services.Helpers;

namespace NetSuiteAutomation.Services
{
    public class ImportFormaService
    {
        private readonly string _logFolder = @"C:\ADSK-Automation\Logs";
        private readonly string _downloadFolder = @"C:\ADSK-Automation\FormaDownloads";
        private readonly string _databasePath = @"C:\ADSK-Automation\Autodesk Transaction Process.accdb";
        private readonly FormaAccessImporter _accessImporter;

        public ImportFormaService(FormaAccessImporter accessImporter)
        {
            _accessImporter = accessImporter;
        }

        public async Task<bool> ImportFormaDataAsync()
        {
            string logFilePath = Path.Combine(_logFolder, "FormaImportIssues.txt");
            string compressedFilePath = Path.Combine(_downloadFolder, "all_transaction_fees.csv.gz");
            string csvFilePath = Path.Combine(_downloadFolder, "all_transaction_fees.csv");

            try
            {
                if (File.Exists(logFilePath))
                    File.Delete(logFilePath);

                Directory.CreateDirectory(_downloadFolder);

                string keyId = "648506c0befa4af6b44d2c1961788ccf";
                string secretKey = "12949c76244262edf1e13e11f2defca8455a7056457e08e1a70aa715e757b079";
                string report = "all_transaction_fees";
                string csn = "0070000545";
                string baseURL = "https://1130-api.forma.ai/report-service/";

                using (HttpClient client = new HttpClient())
                {
                    client.BaseAddress = new Uri(baseURL);
                    client.DefaultRequestHeaders.TryAddWithoutValidation("Authorization", $"KEY_ID={keyId} SECRET_KEY={secretKey}");

                    string requestUrl = $"{baseURL}?CSN={csn}&Report={report}";
                    HttpResponseMessage response = await client.GetAsync(requestUrl);
                    string responseBody = await response.Content.ReadAsStringAsync();

                    Log(logFilePath, $"Response: {response.StatusCode}: {responseBody}");

                    if (!response.IsSuccessStatusCode)
                        throw new Exception("Failed to get Forma report: " + response.StatusCode);

                    dynamic jsonResponse = Newtonsoft.Json.JsonConvert.DeserializeObject(responseBody);
                    string downloadUrl = jsonResponse?.download_url?.ToString();

                    if (string.IsNullOrEmpty(downloadUrl))
                        throw new Exception("Download URL missing in Forma API response.");

                    using (HttpClient downloadClient = new HttpClient())
                    {
                        HttpResponseMessage downloadResponse = await downloadClient.GetAsync(downloadUrl);

                        if (!downloadResponse.IsSuccessStatusCode)
                            throw new Exception("Failed to download CSV: " + downloadResponse.StatusCode);

                        byte[] fileBytes = await downloadResponse.Content.ReadAsByteArrayAsync();
                        File.WriteAllBytes(compressedFilePath, fileBytes);

                        // Decompress
                        using (FileStream original = new FileStream(compressedFilePath, FileMode.Open))
                        using (FileStream decompressed = new FileStream(csvFilePath, FileMode.Create))
                        using (System.IO.Compression.GZipStream unzip = new System.IO.Compression.GZipStream(original, System.IO.Compression.CompressionMode.Decompress))
                        {
                            unzip.CopyTo(decompressed);
                        }

                        Log(logFilePath, $"Decompressed file saved to {csvFilePath}");

                        // Import to Access
                        bool importSuccess = _accessImporter.ImportFormaCsvToAccess(csvFilePath, _databasePath);

                        if (importSuccess)
                        {
                            Log(logFilePath, "✅ Forma data successfully imported into Access.");
                            return true;
                        }
                        else
                        {
                            Log(logFilePath, "❌ Forma import failed during Access insert.");
                            return false;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Log(logFilePath, $"❌ Exception during Forma import: {ex.Message}");
                return false;
            }
        }

        private void Log(string logPath, string message)
        {
            Directory.CreateDirectory(_logFolder);
            File.AppendAllText(logPath, $"{DateTime.Now} - {message}{Environment.NewLine}");

            // Also write to general log
            string generalLog = Path.Combine(_logFolder, "GeneralLog.txt");
            File.AppendAllText(generalLog, $"{DateTime.Now} - {message}{Environment.NewLine}");
        }
    }
}
