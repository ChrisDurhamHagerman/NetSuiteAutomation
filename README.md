# NetSuiteAutomation

What it is (purpose & shape)
An ASP.NET (net8.0) web app that moves data between NetSuite ↔ Access DB ↔ external services (Autodesk/Forma/CreditSafe). It exposes API endpoints (under Controllers/) for NetSuite/SuiteScripts to call, plus simple Razor Pages (under Pages/) so humans can kick off steps, download files, and view logs. Configuration lives in appsettings.json*; static assets (Bootstrap/jQuery) are under wwwroot/. Builds and IIS-ready publish artifacts show up in bin/Release/net8.0/publish (with web.config).

How users & scripts interact

•	Controllers

    o	CsvImportController.cs – endpoints that receive saved-search data from NetSuite and serve files back (e.g., NetsuiteImportData.csv) to NetSuite scheduled imports.
        
    o	ReportsController.cs – lightweight endpoints for downloading/viewing output files and health checks (and/or simple report endpoints).
        
    o	CreditSafeController.cs – endpoints that orchestrate the CreditSafe pipeline (getting NetSuite exports, producing CreditSafeImport.csv, etc.).

•	Pages

      o	CsvImport/View – manual view of recent CSV import activity.
      
      o	FormaImport/FormaDownload – download the latest Forma file(s).
      
      o	CreditSafeImport – run/monitor the CreditSafe flow.
      
      o	Logs/Index – shows text logs written by the services.
      
      o	Index/Privacy/Error – standard site chrome.

What does the heavy lifting (services & helpers)

•	Services

      o	AccessImportService.cs – imports NetSuite CSV/json into Access (cleans data, maps columns, writes to [Autodesk Data-IMPORT]).
      
      o	ImportFormaService.cs – downloads/decompresses the Forma .csv.gz into C:\ADSK-Automation\FormaDownloads.
      
      o	FormaAccessImporter.cs – maps Forma columns, converts types, filters rows, and inserts into Access.
      
      o	AccessMacroService.cs – runs Access macros/exports tables (e.g., builds NetsuiteImportData.csv) and post-processes output.
      
      o	LogService.cs – writes step-by-step logs to C:\ADSK-Automation\Logs that the Pages/Logs screen displays.
  
  •	Services/Helpers
  
    o	CsvReader.cs & CsvCleaner.cs – utility routines to parse CSVs, normalize values (e.g., commas in names, midnight timestamp cleanup), and hand clean rows to the services.

Typical flows (end-to-end)

  1.	NetSuite Saved Search → API → Access → CSV back to NetSuite
  SuiteScript posts data to CsvImportController → AccessImportService loads it → AccessMacroService runs macros & exports NetsuiteImportData.csv → NetSuite scheduled script downloads that file and triggers a Saved CSV Import.
  
  2.	Forma pipeline (optional or chained)
  ImportFormaService fetches Forma report → FormaAccessImporter loads it into Access alongside NetSuite data → AccessMacroService produces updated outputs → available to download from Pages/FormaImport or by API.
  
  4.	CreditSafe pipeline
  CreditSafeController coordinates pulling NetSuite exports, running the Access query (e.g., Process Query 09…), producing CreditSafeImport.csv, which NetSuite then imports via a Saved CSV Import.

