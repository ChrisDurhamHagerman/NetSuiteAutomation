using NetSuiteAutomation.Services;

var builder = WebApplication.CreateBuilder(args);

// ✅ Supports API Controllers
builder.Services.AddControllers();

// ✅ Supports Razor Pages (your /Pages folder)
builder.Services.AddRazorPages();

builder.Services.AddHttpsRedirection(options =>
{
    options.HttpsPort = 44306;
});

builder.Services.AddSingleton<AccessImportService>();
builder.Services.AddSingleton<ImportFormaService>();
builder.Services.AddSingleton<FormaAccessImporter>();
builder.Services.AddSingleton<LogService>();
builder.Services.AddScoped<AccessMacroService>();


var app = builder.Build();

// Configure the HTTP request pipeline.
if (!app.Environment.IsDevelopment())
{
    // ✅ Point to Razor Pages error page, not MVC view
    app.UseExceptionHandler("/Error");
    app.UseHsts();
}

app.UseHttpsRedirection();
app.UseStaticFiles();

app.UseRouting();
app.UseEndpoints(endpoints =>
{
    endpoints.MapControllers();
});

app.UseAuthorization();

// ✅ Razor Pages (like /Index, /CsvImport/View)
app.MapRazorPages();

// ✅ API Controllers (like /api/CsvImport/import)
app.MapControllers();

app.Run();
