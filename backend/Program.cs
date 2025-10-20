using System;
using System.IO;
using System.Text.Json;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

var builder = WebApplication.CreateBuilder(args);

// Add CORS for both local development and production
builder.Services.AddCors(options =>
{
    options.AddDefaultPolicy(policy =>
    {
        policy.WithOrigins(
            "https://shorekoya.github.io",  // GitHub Pages
            "http://localhost:5500",        // Live Server
            "http://127.0.0.1:5500",        // Live Server alternative
            "http://localhost:3000"         // Local development
        )
        .AllowAnyHeader()
        .AllowAnyMethod();
    });
});

var app = builder.Build();

app.UseCors();

// API endpoint for generating reports
app.MapPost("/api/generate-report", async (HttpContext context) =>
{
    try
    {
        // Read the raw request body
        using StreamReader reader = new StreamReader(context.Request.Body);
        string requestBody = await reader.ReadToEndAsync();
        
        Console.WriteLine($"Received request: {requestBody}");

        if (string.IsNullOrEmpty(requestBody))
        {
            return Results.BadRequest("Request body is empty");
        }

        // Parse the JSON manually to see what's happening
        try
        {
            var request = JsonSerializer.Deserialize<ReportRequest>(requestBody, new JsonSerializerOptions
            {
                PropertyNameCaseInsensitive = true
            });

            if (request == null)
            {
                return Results.BadRequest("Invalid JSON format");
            }

            // Validate required fields
            if (string.IsNullOrWhiteSpace(request.ClientName))
            {
                Console.WriteLine("ClientName is missing or empty");
                return Results.BadRequest("Client name is required");
            }

            if (string.IsNullOrWhiteSpace(request.ReportType))
            {
                Console.WriteLine("ReportType is missing or empty");
                return Results.BadRequest("Report type is required");
            }

            Console.WriteLine($"Processing: {request.ClientName}, {request.ReportType}, {request.ReportYear}");

            // Generate the report
            string filePath = ReportGenerator.CreateSimpleReport(
                request.ClientName,
                request.ReportType,
                request.ReportYear
            );

            // Return success response
            return Results.Ok(new
            {
                success = true,
                message = "Report generated successfully",
                filePath = filePath,
                fileName = Path.GetFileName(filePath)
            });
        }
        catch (JsonException jsonEx)
        {
            Console.WriteLine($"JSON parsing error: {jsonEx.Message}");
            return Results.BadRequest($"Invalid JSON format: {jsonEx.Message}");
        }
    }
    catch (Exception ex)
    {
        Console.WriteLine($"General error: {ex.Message}");
        return Results.Problem($"Error generating report: {ex.Message}");
    }
});

// File download endpoint
app.MapGet("/api/download/{fileName}", (string fileName) =>
{
    if (string.IsNullOrEmpty(fileName))
    {
        return Results.BadRequest("File name is required");
    }

    string filePath = Path.Combine(Directory.GetCurrentDirectory(), fileName);
    
    if (!File.Exists(filePath))
    {
        return Results.NotFound("File not found");
    }

    return Results.File(filePath, 
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document", 
        fileName);
});

// Health check endpoint
app.MapGet("/", () => "Financial Report Generator API is running!");

app.Run();

// Your existing ReportGenerator class
public class ReportGenerator
{
    public static string CreateSimpleReport(string clientName, string reportType, int reportYear = 2024)
    {
        // Generate unique filename (sanitize for file system)
        string sanitizedClientName = clientName.Replace(" ", "_").Replace("\\", "").Replace("/", "").Replace(":", "");
        string sanitizedReportType = reportType.Replace(" ", "_").Replace("\\", "").Replace("/", "").Replace(":", "");
        
        string fileName = $"FinancialReport_{sanitizedClientName}_{sanitizedReportType}_{DateTime.Now:yyyyMMdd_HHmmss}.docx";
        string filePath = Path.Combine(Directory.GetCurrentDirectory(), fileName);

        // Create a new Word document
        using (WordprocessingDocument wordDocument = 
            WordprocessingDocument.Create(filePath, WordprocessingDocumentType.Document))
        {
            // Add a main document part
            MainDocumentPart mainPart = wordDocument.AddMainDocumentPart();
            mainPart.Document = new Document();
            Body body = mainPart.Document.AppendChild(new Body());

            // Create and format the title paragraph
            Paragraph titlePara = body.AppendChild(new Paragraph());
            Run titleRun = titlePara.AppendChild(new Run());
            
            // Apply formatting to title
            RunProperties titleProps = titleRun.AppendChild(new RunProperties());
            titleProps.AppendChild(new Bold());
            titleProps.AppendChild(new FontSize() { Val = "32" });
            titleProps.AppendChild(new Color() { Val = "2563EB" });
            
            titleRun.AppendChild(new Text($"Financial Report: {reportType}"));
            
            // Add spacing after title
            ParagraphProperties titleParaProps = titlePara.AppendChild(new ParagraphProperties());
            titleParaProps.AppendChild(new SpacingBetweenLines() { After = "400" });

            // Create client information paragraph
            Paragraph clientPara = body.AppendChild(new Paragraph());
            Run clientRun = clientPara.AppendChild(new Run());
            
            RunProperties clientProps = clientRun.AppendChild(new RunProperties());
            clientProps.AppendChild(new FontSize() { Val = "24" });
            
            clientRun.AppendChild(new Text($"Client: {clientName}"));

            // Create year information paragraph
            Paragraph yearPara = body.AppendChild(new Paragraph());
            Run yearRun = yearPara.AppendChild(new Run());
            
            RunProperties yearProps = yearRun.AppendChild(new RunProperties());
            yearProps.AppendChild(new FontSize() { Val = "24" });
            
            yearRun.AppendChild(new Text($"Reporting Year: {reportYear}"));

            // Add a separator line
            Paragraph separatorPara = body.AppendChild(new Paragraph());
            ParagraphProperties separatorParaProps = separatorPara.AppendChild(new ParagraphProperties());
            separatorParaProps.AppendChild(new ParagraphBorders(
                new BottomBorder() 
                { 
                    Val = BorderValues.Single, 
                    Size = 12, 
                    Color = "E2E8F0" 
                }
            ));
            separatorParaProps.AppendChild(new SpacingBetweenLines() { After = "200", Before = "200" });

            // Add metadata paragraph
            Paragraph metaPara = body.AppendChild(new Paragraph());
            Run metaRun = metaPara.AppendChild(new Run());
            
            RunProperties metaProps = metaRun.AppendChild(new RunProperties());
            metaProps.AppendChild(new FontSize() { Val = "20" });
            metaProps.AppendChild(new Color() { Val = "64748B" });
            metaProps.AppendChild(new Italic());
            
            metaRun.AppendChild(new Text($"Generated on: {DateTime.Now:MMMM dd, yyyy 'at' hh:mm tt}"));

            // Save the document
            mainPart.Document.Save();
        }

        Console.WriteLine($"Report generated successfully: {filePath}");
        return filePath;
    }
}

// Request model
public class ReportRequest
{
    public string ReportType { get; set; } = string.Empty;
    public int ReportYear { get; set; } = DateTime.Now.Year;
    public string ClientName { get; set; } = string.Empty;
    public string Timestamp { get; set; } = string.Empty;
    public string RequestId { get; set; } = string.Empty;
}