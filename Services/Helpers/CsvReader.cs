using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace NetSuiteAutomation.Services.Helpers
{
    public class CsvReader : IDisposable
    {
        private StreamReader _reader;

        public CsvReader(string filePath)
        {
            _reader = new StreamReader(filePath);
        }

        public IEnumerable<string[]> RowEnumerator
        {
            get
            {
                string line;
                while ((line = _reader.ReadLine()) != null)
                {
                    yield return ParseCsvLine(line);
                }
            }
        }

        private string[] ParseCsvLine(string line)
        {
            var fields = new List<string>();
            var current = new StringBuilder();
            bool inQuotes = false;

            for (int i = 0; i < line.Length; i++)
            {
                char c = line[i];

                if (inQuotes)
                {
                    if (c == '"')
                    {
                        if (i + 1 < line.Length && line[i + 1] == '"')
                        {
                            current.Append('"');
                            i++;
                        }
                        else
                        {
                            inQuotes = false;
                        }
                    }
                    else
                    {
                        current.Append(c);
                    }
                }
                else
                {
                    if (c == '"')
                    {
                        inQuotes = true;
                    }
                    else if (c == ',')
                    {
                        fields.Add(current.ToString());
                        current.Clear();
                    }
                    else
                    {
                        current.Append(c);
                    }
                }
            }

            fields.Add(current.ToString());
            return fields.ToArray();
        }

        public void Dispose()
        {
            _reader?.Dispose();
        }
    }
}
