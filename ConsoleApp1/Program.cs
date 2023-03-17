using System;
using System.IO;
using System.Reflection;
using DevExpress.XtraPrinting;
using DevExpress.XtraRichEdit;
using DevExpress.XtraRichEdit.API.Native;

namespace ConsoleApp1
{
    internal class Program
    {
        static void Main(string[] args)
        {
            // Load the document from resources
            var docxDocument = GetEmbeddedResourceFile("ConsoleApp1.document2.docx", null);

            // Convert to PDF before mail merge, this is fine in Adobe Reader
            DocxToPdf(docxDocument, @"document.pdf");

            // Run mail merge
            var docxMailMergedDocument = MailMergeDocument(docxDocument, new[] { new {Field1 = "Hello", Field2 = "world!" } });

            // Convert to PDF the merged document -> Adobe Reader complains!
            DocxToPdf(docxMailMergedDocument, @"mailMergedDocument.pdf");

            Console.WriteLine("Done!");
        }

        private static byte[] GetEmbeddedResourceFile(string embeddedResourceName, Assembly containingAssembly)
        {
            if (containingAssembly == null)
                containingAssembly = Assembly.GetExecutingAssembly();

            using (Stream stream = containingAssembly.GetManifestResourceStream(embeddedResourceName))
            using (MemoryStream reader = new MemoryStream())
            {
                if (stream == null)
                    throw new Exception(String.Format("Resource {0} not found!", embeddedResourceName));

                stream.CopyTo(reader);
                return reader.ToArray();
            }
        }

        private static byte[] MailMergeDocument(byte[] documentBytes, object dataSource)
        {
            using (RichEditDocumentServer docServer = new RichEditDocumentServer())
            {
                docServer.LoadDocument(documentBytes, DocumentFormat.OpenXml);

                MailMergeOptions options = docServer.CreateMailMergeOptions();
                options.DataSource = dataSource;

                byte[] output;
                using (MemoryStream s = new MemoryStream())
                {
                    docServer.MailMerge(options, s, DocumentFormat.OpenXml);
                    output = s.ToArray();
                }

                return output;
            }
        }

        private static void DocxToPdf(byte[] docxDocument, string outputFileName)
        {
            using (RichEditDocumentServer wordProcessor = new RichEditDocumentServer())
            {
                // Load a DOCX document.
                wordProcessor.LoadDocument(docxDocument, DocumentFormat.OpenXml);

                // Specify PDF export options.
                PdfExportOptions options = new PdfExportOptions();

                // Export the document to PDF.
                wordProcessor.ExportToPdf(outputFileName, options);
            }
        }
    }
}