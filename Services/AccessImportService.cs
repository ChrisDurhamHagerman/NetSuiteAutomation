using NetSuiteAutomation.Services.Helpers;
using System;
using System.Data.OleDb;
using System.IO;

namespace NetSuiteAutomation.Services
{
    public class AccessImportService
    {
        private readonly string _logFolder = @"C:\ADSK-Automation\Logs";
        private readonly string _databasePath = @"C:\ADSK-Automation\Autodesk Transaction Process.accdb";

        public bool ImportCsvToAccess(string csvFilePath)
        {
            string logFilePath = Path.Combine(_logFolder, "NetsuiteImportIssues.txt");
            int lineNumber = 0;
            int errorCount = 0;
            string truncateNetsuiteTable = "DELETE FROM [dbo_autodesk_Sales_Orders]";

            if (File.Exists(logFilePath))
            {
                File.Delete(logFilePath);
            }

            try
            {
                string connectionString = $@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={_databasePath};Persist Security Info=False;";
                using (OleDbConnection connection = new OleDbConnection(connectionString))
                {
                    connection.Open();
                    Log("Connected to Access DB");

                    using (OleDbCommand truncateCommand = new OleDbCommand(truncateNetsuiteTable, connection))
                    {
                        truncateCommand.ExecuteNonQuery();
                        Log("Truncated existing Access table data.");
                    }

                    using (CsvReader reader = new CsvReader(csvFilePath))
                    {
                        bool isFirstLine = true;
                        foreach (string[] values in reader.RowEnumerator)
                        {
                            lineNumber++;

                            if (isFirstLine)
                            {
                                isFirstLine = false;
                                continue;
                            }

                            try
                            {
                                const int customerIndex = 1;
                                if (values[customerIndex].Contains(","))
                                    values[customerIndex] = values[customerIndex].Replace(",", " ");

                                double unitCustomerPrice = 0.0;
                                if (!string.IsNullOrEmpty(values[9]))
                                {
                                    if (!double.TryParse(values[9], out unitCustomerPrice))
                                    {
                                        throw new Exception($"Invalid double format on line {lineNumber}, Total Billing (Target Currency): {values[9]}");
                                    }
                                }

                                string insertCommandText = @"
                                    INSERT INTO [dbo_autodesk_Sales_Orders] (
                                        [InternalID], [Customer], [SalesOrder], [AutodeskQuoteNbr], 
                                        [LineID], [AutodeskLineID], [Item], [Item_value], [Quantity], 
                                        [Unit_Customer_Price], [AD_Start_Date], [AD_End_Date]
                                    ) 
                                    VALUES (
                                        @internalID, @customer, @salesOrder, @autodeskQuoteNbr, 
                                        @lineID, @autodeskLineID, @item, @itemValue, @quantity, 
                                        @unitCustomerPrice, @adStartDate, @adEndDate
                                    )";

                                using (OleDbCommand command = new OleDbCommand(insertCommandText, connection))
                                {
                                    command.Parameters.AddWithValue("@internalID", values[0]);
                                    command.Parameters.AddWithValue("@customer", values[1]);
                                    command.Parameters.AddWithValue("@salesOrder", values[2]);
                                    command.Parameters.AddWithValue("@autodeskQuoteNbr", values[3]);
                                    command.Parameters.AddWithValue("@lineID", values[4]);
                                    command.Parameters.AddWithValue("@autodeskLineID", values[5]);
                                    command.Parameters.AddWithValue("@item", values[6]);
                                    command.Parameters.AddWithValue("@itemValue", values[6]);
                                    command.Parameters.AddWithValue("@quantity", values[7]);
                                    command.Parameters.AddWithValue("@unitCustomerPrice", unitCustomerPrice);

                                    if (!string.IsNullOrEmpty(values[10]) && DateTime.TryParse(values[10], out DateTime adStart))
                                        command.Parameters.AddWithValue("@adStartDate", adStart);
                                    else
                                    {
                                        LogError(logFilePath, $"Warning on line {lineNumber}: Invalid AD Start Date format. Value: {values[10]}");
                                        command.Parameters.AddWithValue("@adStartDate", DBNull.Value);
                                    }

                                    if (!string.IsNullOrEmpty(values[11]) && DateTime.TryParse(values[11], out DateTime adEnd))
                                        command.Parameters.AddWithValue("@adEndDate", adEnd);
                                    else
                                    {
                                        LogError(logFilePath, $"Warning on line {lineNumber}: Invalid AD End Date format. Value: {values[11]}");
                                        command.Parameters.AddWithValue("@adEndDate", DBNull.Value);
                                    }

                                    command.ExecuteNonQuery();
                                }
                            }
                            catch (Exception ex)
                            {
                                errorCount++;
                                LogError(logFilePath, $"Error on line {lineNumber}: {ex.Message}. Line content: {string.Join(",", values)}");
                            }
                        }
                    }
                }

                if (errorCount > 0)
                {
                    Log($"Import completed with {errorCount} error(s).");
                }
                else
                {
                    Log("CSV imported successfully with no errors.");
                }

                return errorCount == 0;
            }
            catch (Exception ex)
            {
                LogError(logFilePath, $"Overall error during CSV import: {ex.Message}");
                Log($"Critical error: {ex.Message}");
                return false;
            }
        }

        private void LogError(string logPath, string message)
        {
            Directory.CreateDirectory(_logFolder);
            File.AppendAllText(logPath, $"{DateTime.Now} - {message}{Environment.NewLine}");
        }

        private void Log(string message)
        {
            string generalLogPath = Path.Combine(_logFolder, "GeneralLog.txt");
            Directory.CreateDirectory(_logFolder);
            File.AppendAllText(generalLogPath, $"{DateTime.Now} - {message}{Environment.NewLine}");
        }
    }
}
