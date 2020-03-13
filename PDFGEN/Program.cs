﻿using MigraDoc.Rendering;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PDFGEN
{
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create an invoice form with the sample invoice data.
                var invoice = new InvoiceForm("invoice.xml");

                // Create the document using MigraDoc.
                var document = invoice.CreateDocument();
                document.UseCmykColor = true;

#if DEBUG
                // For debugging only...
                MigraDoc.DocumentObjectModel.IO.DdlWriter.WriteToFile(document, "MigraDoc.mdddl");
                var document2 = MigraDoc.DocumentObjectModel.IO.DdlReader.DocumentFromFile("MigraDoc.mdddl");
                //document = document2;
                // With PDFsharp 1.50 beta 3 there is a known problem: the blank before "by" gets lost while persisting as MDDDL.
#endif

                // Create a renderer for PDF that uses Unicode font encoding.
                var pdfRenderer = new PdfDocumentRenderer(true);

                // Set the MigraDoc document.
                pdfRenderer.Document = document;

                // Create the PDF document.
                pdfRenderer.RenderDocument();

                // Save the PDF document...
                var filename = "Invoice.pdf";
#if DEBUG
                // I don't want to close the document constantly...
                filename = "Invoice-" + Guid.NewGuid().ToString("N").ToUpper() + ".pdf";
#endif
                pdfRenderer.Save(filename);
                // ...and start a viewer.
                Process.Start(filename);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                Console.ReadLine();
            }
        }
    }
 }

