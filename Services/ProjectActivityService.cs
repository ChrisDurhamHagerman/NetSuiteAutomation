using System;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Text;

public class ProjectActivityService
{
    private readonly string _databasePath = @"C:\ADSK-Automation\Project Activities.accdb";
    private readonly string _logFolder = @"C:\ADSK-Automation\Logs";

    public bool ImportWorkflowCsvToAccess(string csvPath)
    {
        string logPath = Path.Combine(_logFolder, "WorkflowImportIssues.txt");
        string fileName = Path.GetFileName(csvPath);
        string csvFolder = Path.GetDirectoryName(csvPath) ?? "";

        try
        {
            Directory.CreateDirectory(_logFolder);

            if (!File.Exists(csvPath))
            {
                File.AppendAllText(logPath, $"[{DateTime.Now}] ERROR: CSV file not found: {csvPath}{Environment.NewLine}");
                return false;
            }

            // Log header (no mutation)
            using (var sr = new StreamReader(csvPath, Encoding.Default, true))
            {
                string? header = sr.ReadLine();
                File.AppendAllText(logPath, $"[{DateTime.Now}] INFO: CSV Header: {header}{Environment.NewLine}");
            }

            // DO NOT REWRITE THE FILE — it is already in Access-safe format

            // Write schema.ini matching our format
            WriteSchemaIni(csvFolder, fileName, logPath);

            // Preflight: confirm ACE sees expected columns
            PreflightLogTextDriverHeaders(csvFolder, fileName, logPath);

            // Import into Access
            string connStr = $@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={_databasePath};Persist Security Info=False;";
            using (OleDbConnection connection = new OleDbConnection(connStr))
            {
                connection.Open();

                string importQuery = $@"
INSERT INTO [Project Activities] 
    ([Internal ID], [Sales Rep], [Company/Project], [Created By], [Subject], [Date], [Comment])
SELECT 
    [Internal ID], [Sales Rep], [Company/Project], [Created By], [Subject], [Date], [Comment]
FROM [Text;FMT=Delimited;HDR=YES;Database={csvFolder};].[{fileName}]";

                File.AppendAllText(logPath, $"[{DateTime.Now}] INFO: Import SQL:{Environment.NewLine}{importQuery}{Environment.NewLine}");

                using (var cmd = new OleDbCommand(importQuery, connection))
                {
                    int rows = cmd.ExecuteNonQuery();
                    File.AppendAllText(logPath, $"[{DateTime.Now}] SUCCESS: Inserted {rows} records from {fileName}{Environment.NewLine}");
                }
            }

            return true;
        }
        catch (Exception ex)
        {
            File.AppendAllText(logPath, $"[{DateTime.Now}] ERROR: {ex.Message}{Environment.NewLine}");
            return false;
        }
    }

    private static void WriteSchemaIni(string folderPath, string fileName, string logPath)
    {
        string schemaPath = Path.Combine(folderPath, "schema.ini");

        string schema =
$@"[{fileName}]
Format=CSVDelimited
ColNameHeader=True
CharacterSet=ANSI
MaxScanRows=0
Col1=""Internal ID"" Long
Col2=""Sales Rep"" Text Width 255
Col3=""Company/Project"" Text Width 255
Col4=""Created By"" Text Width 255
Col5=""Subject"" Text Width 255
Col6=""Date"" DateTime
Col7=""Comment"" Memo
";

        try
        {
            File.WriteAllText(schemaPath, schema.Replace("\n", "\r\n"), Encoding.Default);
            File.AppendAllText(logPath, $"[{DateTime.Now}] INFO: schema.ini content:{Environment.NewLine}{schema}{Environment.NewLine}");
        }
        catch (Exception ex)
        {
            File.AppendAllText(logPath, $"[{DateTime.Now}] WARNING: Could not write schema.ini: {ex.Message}{Environment.NewLine}");
        }
    }

    private static void PreflightLogTextDriverHeaders(string folderPath, string fileName, string logPath)
    {
        try
        {
            string textConn =
                $@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={folderPath};Extended Properties=""text;HDR=Yes;FMT=Delimited"";";

            using (var cn = new OleDbConnection(textConn))
            using (var cmd = new OleDbCommand($"SELECT TOP 1 * FROM [{fileName}]", cn))
            {
                cn.Open();
                using var rdr = cmd.ExecuteReader(CommandBehavior.SchemaOnly);
                if (rdr != null)
                {
                    var sb = new StringBuilder();
                    sb.Append("ACE sees columns: ");
                    for (int i = 0; i < rdr.FieldCount; i++)
                    {
                        if (i > 0) sb.Append(" | ");
                        sb.Append(rdr.GetName(i));
                    }

                    File.AppendAllText(logPath, $"[{DateTime.Now}] INFO: {sb}{Environment.NewLine}");
                }
            }
        }
        catch (Exception ex)
        {
            File.AppendAllText(logPath, $"[{DateTime.Now}] WARNING: Preflight header read failed: {ex.Message}{Environment.NewLine}");
        }
    }
}
