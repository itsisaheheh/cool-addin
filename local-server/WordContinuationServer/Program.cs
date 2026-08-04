using System.Security.Cryptography.X509Certificates;
using Microsoft.AspNetCore.StaticFiles;
using Microsoft.Extensions.FileProviders;

var builder = WebApplication.CreateBuilder(args);

var certificatePath = Path.Combine(
    AppContext.BaseDirectory,
    "certificate",
    "WordContinuationChecker.pfx"
);

if (!File.Exists(certificatePath))
{
    Console.Error.WriteLine($"HTTPS certificate was not found: {certificatePath}");
    return;
}

var certificate = new X509Certificate2(
    certificatePath,
    "WCC-Local-5016!"
);

builder.WebHost.ConfigureKestrel(options =>
{
    options.ListenLocalhost(5016, listenOptions =>
    {
        listenOptions.UseHttps(certificate);
    });
});

var app = builder.Build();

var distPath = Path.Combine(AppContext.BaseDirectory, "dist");

if (!Directory.Exists(distPath))
{
    Console.Error.WriteLine($"The dist folder was not found: {distPath}");
    return;
}

var fileProvider = new PhysicalFileProvider(distPath);
var contentTypeProvider = new FileExtensionContentTypeProvider();
contentTypeProvider.Mappings[".xml"] = "application/xml";

app.UseDefaultFiles(new DefaultFilesOptions
{
    FileProvider = fileProvider
});

app.UseStaticFiles(new StaticFileOptions
{
    FileProvider = fileProvider,
    ContentTypeProvider = contentTypeProvider
});

app.MapGet(
    "/health",
    () => Results.Ok("Word Continuation Checker server is running.")
);

app.Run();