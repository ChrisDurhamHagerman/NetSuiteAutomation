public static class CsvCleaner
{
    public static void CleanCsvFileOfMidnightTime(string filePath, string logFilePath)
    {
        try
        {
            var lines = File.ReadAllLines(filePath);
            for (int i = 0; i < lines.Length; i++)
            {
                lines[i] = lines[i].Replace(" 0:00:00", "");
            }
            File.WriteAllLines(filePath, lines);
            // You can optionally extract logging if needed
            File.AppendAllText(logFilePath, $"{DateTime.Now} - CSV file scrubbed of ' 0:00:00' successfully.\n");
        }
        catch (Exception ex)
        {
            File.AppendAllText(logFilePath, $"{DateTime.Now} - Error scrubbing CSV: {ex.Message}\n");
        }
    }
}
