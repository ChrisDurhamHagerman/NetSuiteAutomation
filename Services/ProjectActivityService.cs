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

        try
        {
            Directory.CreateDirectory(_logFolder);

            if (!File.Exists(csvPath))
            {
                File.AppendAllText(logPath, $"[{DateTime.Now}] ERROR: CSV file not found: {csvPath}{Environment.NewLine}");
                return false;
            }

            // Log header
            using (var sr = new StreamReader(csvPath, Encoding.Default, true))
            {
                string? header = sr.ReadLine();
                File.AppendAllText(logPath, $"[{DateTime.Now}] INFO: CSV Header: {header}{Environment.NewLine}");
            }

            string connStr = $@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={_databasePath};Persist Security Info=False;";

            using (OleDbConnection connection = new OleDbConnection(connStr))
            {
                connection.Open();

                string insertSql = @"
INSERT INTO [Project Activities]
([Internal ID], [Sales Rep], [Company/Project], [Created By], [Subject], [Date], [Comment])
VALUES (?, ?, ?, ?, ?, ?, ?)";

                using var cmd = new OleDbCommand(insertSql, connection);

                int successCount = 0;
                int rowNumber = 1;

                foreach (var line in File.ReadLines(csvPath).Skip(1)) // skip header
                {
                    try
                    {
                        var parts = SplitCsvLine(line);

                        cmd.Parameters.Clear();
                        cmd.Parameters.AddWithValue("@p1", int.Parse(parts[0]));
                        cmd.Parameters.AddWithValue("@p2", parts[1]);
                        cmd.Parameters.AddWithValue("@p3", parts[2]);
                        cmd.Parameters.AddWithValue("@p4", parts[3]);
                        cmd.Parameters.AddWithValue("@p5", parts[4]);
                        cmd.Parameters.AddWithValue("@p6", DateTime.Parse(parts[5]));
                        cmd.Parameters.AddWithValue("@p7", parts[6]);

                        cmd.ExecuteNonQuery();
                        successCount++;
                    }
                    catch (Exception exRow)
                    {
                        File.AppendAllText(logPath, $"[{DateTime.Now}] ERROR on row {rowNumber}: {exRow.Message}{Environment.NewLine}");
                    }

                    rowNumber++;
                }

                File.AppendAllText(logPath, $"[{DateTime.Now}] SUCCESS: Inserted {successCount} records from {csvPath}{Environment.NewLine}");
            }

            return true;
        }
        catch (Exception ex)
        {
            File.AppendAllText(logPath, $"[{DateTime.Now}] ERROR: {ex.Message}{Environment.NewLine}");
            return false;
        }
    }

    private static string[] SplitCsvLine(string line)
    {
        var result = new List<string>();
        var current = new StringBuilder();
        bool inQuotes = false;

        foreach (char c in line)
        {
            if (c == '"')
            {
                inQuotes = !inQuotes;
                continue;
            }

            if (c == ',' && !inQuotes)
            {
                result.Add(current.ToString());
                current.Clear();
            }
            else
            {
                current.Append(c);
            }
        }

        result.Add(current.ToString());
        return result.ToArray();
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
