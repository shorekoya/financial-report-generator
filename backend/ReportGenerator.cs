using System;
using System.IO;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace FinancialReportGenerator
{
    /// <summary>
    /// Server-side financial report generator using OpenXML SDK
    /// NuGet Package Required: DocumentFormat.OpenXml (v2.20.0 or later)
    /// Namespaces: DocumentFormat.OpenXml, DocumentFormat.OpenXml.Packaging, DocumentFormat.OpenXml.Wordprocessing
    /// </summary>
    public class ReportGenerator
    {
        /// <summary>
        /// Creates a simple financial report Word document
        /// </summary>
        /// <param name="clientName">Name or ID of the client</param>
        /// <param name="reportType">Type of report (P&L, Balance Sheet, Cash Flow)</param>
        /// <param name="reportYear">The fiscal year for the report</param>
        /// <returns>Path to the generated document</returns>
        public static string CreateSimpleReport(string clientName, string reportType, int reportYear = 2024)
        {
            // Generate unique filename
            string fileName = $"GeneratedReport_{DateTime.Now:yyyyMMdd_HHmmss}.docx";
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
                titleProps.AppendChild(new FontSize() { Val = "32" }); // 16pt font
                titleProps.AppendChild(new Color() { Val = "2563EB" }); // Blue color
                
                titleRun.AppendChild(new Text($"Financial Report: {reportType}"));
                
                // Add spacing after title
                ParagraphProperties titleParaProps = titlePara.AppendChild(new ParagraphProperties());
                titleParaProps.AppendChild(new SpacingBetweenLines() { After = "400" });

                // Create client information paragraph
                Paragraph clientPara = body.AppendChild(new Paragraph());
                Run clientRun = clientPara.AppendChild(new Run());
                
                RunProperties clientProps = clientRun.AppendChild(new RunProperties());
                clientProps.AppendChild(new FontSize() { Val = "24" }); // 12pt font
                
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
                metaProps.AppendChild(new FontSize() { Val = "20" }); // 10pt font
                metaProps.AppendChild(new Color() { Val = "64748B" }); // Gray color
                metaProps.AppendChild(new Italic());
                
                metaRun.AppendChild(new Text($"Generated on: {DateTime.Now:MMMM dd, yyyy 'at' hh:mm tt}"));

                // Save the document
                mainPart.Document.Save();
            }

            Console.WriteLine($"Report generated successfully: {filePath}");
            return filePath;
        }

        /// <summary>
        /// Complete API endpoint implementation
        /// </summary>
        public static void GenerateReportFromRequest(ReportRequest request)
        {
            try
            {
                // Validate request
                if (string.IsNullOrWhiteSpace(request.ClientName))
                    throw new ArgumentException("Client name is required");

                if (string.IsNullOrWhiteSpace(request.ReportType))
                    throw new ArgumentException("Report type is required");

                // Generate the report
                string filePath = CreateSimpleReport(
                    request.ClientName, 
                    request.ReportType, 
                    request.ReportYear
                );

                Console.WriteLine($"Report created for {request.ClientName}: {filePath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error generating report: {ex.Message}");
                throw;
            }
        }
    }

    /// <summary>
    /// Request model matching the frontend JSON payload
    /// </summary>
    public class ReportRequest
    {
        public string ReportType { get; set; }
        public int ReportYear { get; set; }
        public string ClientName { get; set; }
        public string Timestamp { get; set; }
        public string RequestId { get; set; }

        // Add constructor to initialize properties
        public ReportRequest()
        {
            ReportType = string.Empty;
            ClientName = string.Empty;
            Timestamp = string.Empty;
            RequestId = string.Empty;
            ReportYear = DateTime.Now.Year;
        }
    }

    /// <summary>
    /// Example usage and testing
    /// </summary>
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Financial Report Generator - Backend Service\n");
            Console.WriteLine("=".PadRight(50, '=') + "\n");

            // Test Case 1: P&L Report
            Console.WriteLine("Test 1: Generating P&L Report...");
            ReportGenerator.CreateSimpleReport("Acme Corporation", "P&L", 2024);  // ✅ Fixed
            Console.WriteLine();

            // Test Case 2: Balance Sheet
            Console.WriteLine("Test 2: Generating Balance Sheet...");
            ReportGenerator.CreateSimpleReport("TechStart Industries", "Balance Sheet", 2024);  // ✅ Fixed
            Console.WriteLine();

            // Test Case 3: Cash Flow with API model
            Console.WriteLine("Test 3: Generating Cash Flow via API model...");
            var request = new ReportRequest
            {
                ClientName = "Global Finance Ltd",
                ReportType = "Cash Flow",
                ReportYear = 2024,
                Timestamp = DateTime.Now.ToString("o"),
                RequestId = "REQ-12345"
            };
            ReportGenerator.GenerateReportFromRequest(request);
            Console.WriteLine();

            Console.WriteLine("=".PadRight(50, '='));
            Console.WriteLine("All tests completed successfully!");
            Console.ReadLine();
        }
    }
}