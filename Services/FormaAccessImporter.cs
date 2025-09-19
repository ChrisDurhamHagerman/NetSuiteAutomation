using NetSuiteAutomation.Services.Helpers;
using System;
using System.Data.OleDb;
using System.IO;

namespace NetSuiteAutomation.Services
{
    public class FormaAccessImporter
    {
        private readonly string _logFolder = @"C:\ADSK-Automation\Logs";
        private readonly string _databasePath = @"C:\ADSK-Automation\Autodesk Transaction Process.accdb";

        public bool ImportFormaCsvToAccess(string csvFilePath, string databasePath)
        {
            string logFilePath = Path.Combine(_logFolder, "FormaImportIssues.txt");
            int lineNumber = 0;
            int errorCount = 0;

            if (File.Exists(logFilePath))
                File.Delete(logFilePath);

            try
            {
                string connectionString = $@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={_databasePath};Persist Security Info=False;";
                using (OleDbConnection connection = new OleDbConnection(connectionString))
                {
                    connection.Open();

                    string truncateQuery = "DELETE FROM [Autodesk Data-IMPORT]";
                    using (OleDbCommand truncateCommand = new OleDbCommand(truncateQuery, connection))
                    {
                        truncateCommand.ExecuteNonQuery();
                    }

                    using (CsvReader reader = new CsvReader(csvFilePath))
                    {
                        foreach (string[] values in reader.RowEnumerator)
                        {
                            lineNumber++;
                            if (lineNumber == 1)
                                continue;

                            try
                            {
                                string isMDP = values[1];
                                string isOffline = values[3];

                                if (isMDP == "MDP" || isOffline == "Offline Adjustments" || isOffline == "Adjustment")
                                {
                                    LogError(logFilePath, $"Line {lineNumber} skipped due to program: {isMDP} / {isOffline}");
                                    continue;
                                }

                                if (values[10].Contains(","))
                                    values[10] = values[10].Replace(",", " ");

                                DateTime.TryParse(values[25], out DateTime invoiceDate);
                                DateTime.TryParse(values[24], out DateTime originalInvoiceDate);
                                DateTime.TryParse(values[48], out DateTime contractEndDate);
                                DateTime.TryParse(values[117], out DateTime settlementStartDate);
                                DateTime.TryParse(values[120], out DateTime settlementEndDate);

                                double.TryParse(values[72], out double totalBillingBillingCurrency);
                                double.TryParse(values[83], out double payoutAmountBillingCurrency);
                                int.TryParse(values[178], out int billingSequence);

                                string insertCommandText = @"
                        INSERT INTO [Autodesk Data-IMPORT] (
                            [Program],
                            [Sold-to Customer #],
                            [Sold-to Customer Name],
                            [Quote #],
                            [Quote Line Item Id],
                            [Customer PO #],
                            [Invoice #],
                            [Original Invoice Date],
                            [Invoice Date],
                            [Original Order #],
                            [SAP Product Line],
                            [Product Class],
                            [Product Name],
                            [Contract Term],
                            [Contract End Date],
                            [WWS Offer Type Group Detail],
                            [Total Billing (Target Currency)],
                            [Payout Amount (Billing Currency)],
                            [QTD Payout %],
                            [Original Order Seat],
                            [Quantity Billed],
                            [Settlement Start Date],
                            [Settlement End Date],
                            [Subscription ID],
                            [Reference Subscription ID],
                            [Payment Type],
                            [Billing Sequence Number],
                            [End User Trade (Company) # (CSN)]
                        ) 
                        VALUES (
                            @program, @soldToCustomerNumber, @soldToCustomerName, @quoteNumber, 
                            @quoteLineItemId, @customerPoNumber, @invoiceNumber, @originalInvoiceDate, @invoiceDate, @originalOrderNumber, @sapProductLine, 
                            @productClass, @productName, @contractTerm, @contractEndDate, @wwsOfferTypeGroupDetail,
                            @totalBillingBillingCurrency, @payoutAmountBillingCurrency, @qtdPayout, @originalOrderSeat, @quantityBilled,
                            @settlementStartDate, @settlementEndDate, @subscriptionId, @referenceSubscriptionId, @paymentType, @billingSequence, @EndUserTradeCompanyCSN
                        )";

                                using (OleDbCommand command = new OleDbCommand(insertCommandText, connection))
                                {
                                    // Map the parameters to the corresponding CSV columns
                                    command.Parameters.AddWithValue("@program", values[1]);
                                    command.Parameters.AddWithValue("@soldToCustomerNumber", values[9]);
                                    command.Parameters.AddWithValue("@soldToCustomerName", values[10]); // Customer name
                                    command.Parameters.AddWithValue("@quoteNumber", values[11]);
                                    command.Parameters.AddWithValue("@quoteLineItemId", string.IsNullOrEmpty(values[12]) ? "*" : values[12]);
                                    command.Parameters.AddWithValue("@customerPoNumber", values[19]);
                                    command.Parameters.AddWithValue("@invoiceNumber", values[20]);
                                    command.Parameters.AddWithValue("@originalInvoiceDate", originalInvoiceDate);
                                    command.Parameters.AddWithValue("@invoiceDate", invoiceDate);
                                    command.Parameters.AddWithValue("@originalOrderNumber", values[26]);
                                    command.Parameters.AddWithValue("@sapProductLine", values[33]);
                                    command.Parameters.AddWithValue("@productClass", values[34]);
                                    command.Parameters.AddWithValue("@productName", values[35]);
                                    command.Parameters.AddWithValue("@contractTerm", values[45]);
                                    command.Parameters.AddWithValue("@contractEndDate", contractEndDate);
                                    command.Parameters.AddWithValue("@wwsOfferTypeGroupDetail", values[57]);
                                    command.Parameters.AddWithValue("@totalBillingBillingCurrency", totalBillingBillingCurrency);
                                    command.Parameters.AddWithValue("@payoutAmountBillingCurrency", payoutAmountBillingCurrency);
                                    command.Parameters.AddWithValue("@qtdPayout", values[85]);
                                    command.Parameters.AddWithValue("@originalOrderSeat", values[110]);
                                    command.Parameters.AddWithValue("@quantityBilled", values[113]);
                                    command.Parameters.AddWithValue("@settlementStartDate", settlementStartDate);
                                    command.Parameters.AddWithValue("@settlementEndDate", settlementEndDate);
                                    command.Parameters.AddWithValue("@subscriptionId", values[155]);
                                    command.Parameters.AddWithValue("@referenceSubscriptionId", values[156]);
                                    command.Parameters.AddWithValue("@paymentType", values[176]);
                                    command.Parameters.AddWithValue("@billingSequence", billingSequence);
                                    command.Parameters.AddWithValue("@EndUserTradeCompanyCSN", values[38]);

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
                    LogError(logFilePath, $"⚠️ {errorCount} line(s) had issues during the Forma import.");
                    return false;
                }

                LogError(logFilePath, "✅ Forma subscription data processed successfully.");
                return true;
            }
            catch (Exception ex)
            {
                LogError(logFilePath, $"❌ Overall error during Forma import: {ex.Message}");
                return false;
            }
        }

        private void LogError(string path, string message)
        {
            Directory.CreateDirectory(_logFolder);
            File.AppendAllText(path, DateTime.Now + " - " + message + Environment.NewLine);
        }
    }
}
