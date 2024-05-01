using dss_sharp;
using GeoPerdasAPI.Services;
using GeoPerdasCloud.ProgGeoPerdas.Legacy.Config;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.FileProviders;
using System.Reflection.Metadata;
using System;
using GeoPerdasCloud.ProgGeoPerdas.Legacy.LegacyCode;

var builder = WebApplication.CreateBuilder(args);
using ILoggerFactory loggerFactory =
            LoggerFactory.Create(builder =>
            builder.AddSimpleConsole(options =>
            {
                options.IncludeScopes = true;
                options.SingleLine = true;
                options.TimestampFormat = "HH:mm:ss ";
            }));
var _logger = loggerFactory.CreateLogger<Program>();

// Add services to the container.
// Learn more about configuring Swagger/OpenAPI at https://aka.ms/aspnetcore/swashbuckle
builder.Services.AddEndpointsApiExplorer();
builder.Services.AddSwaggerGen();

builder.Services.AddCors(options =>
{
    options.AddDefaultPolicy(builder =>
    {
        builder.AllowAnyOrigin()
               .AllowAnyHeader()
               .AllowAnyMethod();
    });
});

builder.Services.AddSpaStaticFiles(configuration =>
{
    configuration.RootPath = "frontapp"; // Update the root path to the Angular app folder
});

var app = builder.Build();

// Configure the app to serve the Angular app
app.UseStaticFiles();
app.UseSpaStaticFiles();

// Configure the HTTP request pipeline.
app.UseSwagger();
app.UseSwaggerUI();

app.UseCors();

app.UseRouting();

// Configure the SPA fallback route
app.MapFallbackToFile("index.html", new StaticFileOptions
{
    FileProvider = new PhysicalFileProvider(Path.Combine(Directory.GetCurrentDirectory(), "frontapp")), // Update the root path to the Angular app folder
    RequestPath = ""
});

app.MapPost("/Request/", ([FromBody] FormConfigControls config) =>
{
    try
    {
        
        _logger.LogInformation("Requesição recebida");        
        if (config == null)
        {            
            
            _logger.LogWarning("Requisição inválida");
            
            return Results.BadRequest();
        }
        SequenceRequestService.RequestSequenceMessage(config);
        
        return Results.Ok("Request processed successfully.");
    }
    catch (Exception ex) {        
        _logger.LogError("Erro na requisição",ex);
        throw ex;
    }   
});

app.MapPost("/ExportDSS/", ([FromBody]FormConfigControls config) =>
{
    var memoryStream = ExportDSSService.ExportDss(config);
    return Results.File(memoryStream.ToArray(), "application/zip", "arquivo.zip");
});

app.Run();
