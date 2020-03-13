using MigraDoc.DocumentObjectModel;
using MigraDoc.DocumentObjectModel.Shapes;
using MigraDoc.DocumentObjectModel.Tables;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using System.Xml.XPath;

namespace PDFGEN
{

    /// <summary>
    /// Creates the invoice form.
    /// </summary>
    public class InvoiceForm
    {
        /// <summary>
        /// The MigraDoc document that represents the invoice.
        /// </summary>
        Document _document;

        /// <summary>
        /// The root navigator for the XML document.
        /// </summary>
        readonly XPathNavigator _navigator;

        /// <summary>
        /// The text frame of the MigraDoc document that contains the address.
        /// </summary>
        TextFrame _addressFrame;

        /// <summary>
        /// The table of the MigraDoc document that contains the invoice items.
        /// </summary>
        Table _table, _table2;

        /// <summary>
        /// Initializes a new instance of the class InvoiceForm and opens the specified XML document.
        /// </summary>
        public InvoiceForm(string filename)
        {
            var invoice = new XmlDocument();
            // An XML invoice based on a sample created with Microsoft InfoPath.
            invoice.Load(filename);
            _navigator = invoice.CreateNavigator();
        }

        /// <summary>
        /// Creates the invoice document.
        /// </summary>
        public Document CreateDocument()
        {
            // Create a new MigraDoc document.
            _document = new Document();
            _document.Info.Title = "A sample invoice";
            _document.Info.Subject = "Demonstrates how to create an invoice.";
            _document.Info.Author = "Stefan Lange";

            DefineStyles();

            CreatePage();

            FillContent();

            return _document;
        }

        /// <summary>
        /// Defines the styles used to format the MigraDoc document.
        /// </summary>
        void DefineStyles()
        {
            // Get the predefined style Normal.
            var style = _document.Styles["Normal"];
            // Because all styles are derived from Normal, the next line changes the 
            // font of the whole document. Or, more exactly, it changes the font of
            // all styles and paragraphs that do not redefine the font.
            style.Font.Name = "Segoe UI";

            style = _document.Styles[StyleNames.Header];
            style.ParagraphFormat.AddTabStop("16cm", TabAlignment.Right);

            style = _document.Styles[StyleNames.Footer];
            style.ParagraphFormat.AddTabStop("8cm", TabAlignment.Center);

            // Create a new style called Table based on style Normal.
            style = _document.Styles.AddStyle("Table", "Normal");
            style.Font.Name = "Segoe UI Semilight";
            style.Font.Size = 9;

            // Create a new style called Title based on style Normal.
            style = _document.Styles.AddStyle("Title", "Normal");
            style.Font.Name = "Segoe UI Semibold";
            style.Font.Size = 9;

            // Create a new style called Reference based on style Normal.
            style = _document.Styles.AddStyle("Reference", "Normal");
            style.ParagraphFormat.SpaceBefore = "5mm";
            style.ParagraphFormat.SpaceAfter = "5mm";
            style.ParagraphFormat.TabStops.AddTabStop("16cm", TabAlignment.Right);
        }

        /// <summary>
        /// Creates the static parts of the invoice.
        /// </summary>
        void CreatePage()
        {
            // Each MigraDoc document needs at least one section.
            var section = _document.AddSection();

            // Define the page setup. We use an image in the header, therefore the
            // default top margin is too small for our invoice.
            section.PageSetup = _document.DefaultPageSetup.Clone();
            // We increase the TopMargin to prevent the document body from overlapping the page header.
            // We have an image of 3.5 cm height in the header.
            // The default position for the header is 1.25 cm.
            // We add 0.5 cm spacing between header image and body and get 5.25 cm.
            // Default value is 2.5 cm.
            section.PageSetup.TopMargin = "5.25cm";

            // Create the footer.
            var paragraph = section.Footers.Primary.AddParagraph();
            paragraph.AddText("Yawer and Shazad Awesome Pty Ltd :-)");
            paragraph.Format.Font.Size = 9;
            paragraph.Format.Alignment = ParagraphAlignment.Center;

            // Create the text frame for the address.
            _addressFrame = section.AddTextFrame();
            _addressFrame.Height = "3.0cm";
            _addressFrame.Width = "16.0cm";
            _addressFrame.Left = ShapePosition.Left;
            _addressFrame.RelativeHorizontal = RelativeHorizontal.Margin;
            _addressFrame.Top = "2.0cm";
            _addressFrame.RelativeVertical = RelativeVertical.Page;

            // Show the sender in the address frame.
            paragraph = _addressFrame.AddParagraph("Seriel Number Search Certificate");
            paragraph.Format.Font.Size = 13;
            paragraph.Format.Font.Bold = true;
            paragraph.Format.SpaceAfter = 3;
            paragraph.Format.Font.Italic = true;
            paragraph.Format.Font.Bold = true;
            paragraph.Format.Alignment = ParagraphAlignment.Center;

            // Create the item table.
            _table = section.AddTable();
            _table.Style = "Table";
            _table.Borders.Color = Colors.Transparent;
            _table.Borders.Width = 0.25;
            _table.Borders.Left.Width = 0.5;
            _table.Borders.Right.Width = 0.5;
            _table.Rows.LeftIndent = 0;

            // Before you can add a row, you must define the columns.
            var column = _table.AddColumn("8.0cm");
            column.Format.Alignment = ParagraphAlignment.Center;

            column = _table.AddColumn("8.0cm");
            column.Format.Alignment = ParagraphAlignment.Right;

            // Create the header of the table.
            var row = _table.AddRow();
            row.HeadingFormat = true;
            row.Format.Alignment = ParagraphAlignment.Left;
            row.Format.Font.Bold = true;
            row.Shading.Color = TableBlue;
            row.Format.SpaceBefore = "1.0cm";
            row.Format.Font.Size = 13;
            row.Format.Font.Color = Colors.Navy;
            row.Cells[0].AddParagraph("IDENTIFIER NUMBER");
            row.Cells[0].Format.Font.Bold = false;
            row.Cells[1].AddParagraph("IDENTIFIER TYPE");
            row.Cells[1].Format.Alignment = ParagraphAlignment.Right;
            row = _table.AddRow();
            row.HeadingFormat = true;
            row.Format.Alignment = ParagraphAlignment.Left;
            row.Format.Font.Bold = true;
            row.Shading.Color = TableBlue;
            row.Cells[0].AddParagraph("1234567891234567");
            row.Cells[0].Format.Font.Bold = false;
            row.Cells[1].AddParagraph("VIN");
            row.Cells[1].Format.Alignment = ParagraphAlignment.Right;

            row = _table.AddRow();
            row.HeadingFormat = true;
            row.Format.Alignment = ParagraphAlignment.Left;
            row.Format.Font.Bold = true;
            row.Shading.Color = TableBlue;
            row.Format.SpaceBefore = "0.8cm";
            row.Format.Font.Size = 13;
            row.Format.Font.Color = Colors.Navy;
            row.Cells[0].AddParagraph("VEHICLE TYPE");
            row.Cells[0].Format.Font.Bold = false;
            row.Cells[1].AddParagraph("MAKE");
            row.Cells[1].Format.Alignment = ParagraphAlignment.Right;
            row = _table.AddRow();
            row.HeadingFormat = true;
            row.Format.Alignment = ParagraphAlignment.Left;
            row.Format.Font.Bold = true;
            row.Shading.Color = TableBlue;
            row.Cells[0].AddParagraph("big truck with trailer");
            row.Cells[0].Format.Font.Bold = false;
            row.Cells[1].AddParagraph("Hino");
            row.Cells[1].Format.Alignment = ParagraphAlignment.Right;
            
            row = _table.AddRow();
            row.HeadingFormat = true;
            row.Format.Alignment = ParagraphAlignment.Left;
            row.Format.Font.Bold = true;
            row.Shading.Color = TableBlue;
            row.Format.SpaceBefore = "0.8cm";
            row.Format.Font.Size = 13;
            row.Format.Font.Color = Colors.Navy;
            row.Cells[0].AddParagraph("BODY TYPE");
            row.Cells[0].Format.Font.Bold = false;
            row.Cells[1].AddParagraph("MODEL");
            row.Cells[1].Format.Alignment = ParagraphAlignment.Right;
            row = _table.AddRow();
            row.HeadingFormat = true;
            row.Format.Alignment = ParagraphAlignment.Left;
            row.Format.Font.Bold = true;
            row.Shading.Color = TableBlue;
            row.Cells[0].AddParagraph("double axle dog type  trailer");
            row.Cells[0].Format.Font.Bold = false;
            row.Cells[1].AddParagraph("5681FG");
            row.Cells[1].Format.Alignment = ParagraphAlignment.Right;


            row = _table.AddRow();
            row.HeadingFormat = true;
            row.Format.Alignment = ParagraphAlignment.Left;
            row.Format.Font.Bold = true;
            row.Shading.Color = TableBlue;
            row.Format.SpaceBefore = "0.8cm";
            row.Format.Font.Size = 13;
            row.Format.Font.Color = Colors.Navy;
            row.Cells[0].AddParagraph("COLOUR");
            row.Cells[0].Format.Font.Bold = false;
            row.Cells[1].AddParagraph("ENGINE NUMBER");
            row.Cells[1].Format.Alignment = ParagraphAlignment.Right;
            row = _table.AddRow();
            row.HeadingFormat = true;
            row.Format.Alignment = ParagraphAlignment.Left;
            row.Format.Font.Bold = true;
            row.Shading.Color = TableBlue;
            row.Cells[0].AddParagraph("Pearl white");
            row.Cells[0].Format.Font.Bold = false;
            row.Cells[1].AddParagraph("123121212");
            row.Cells[1].Format.Alignment = ParagraphAlignment.Right;

            row = _table.AddRow();
            row.HeadingFormat = true;
            row.Format.Alignment = ParagraphAlignment.Left;
            row.Format.Font.Bold = true;
            row.Shading.Color = TableBlue;
            row.Format.SpaceBefore = "0.8cm";
            row.Format.Font.Size = 13;
            row.Format.Font.Color = Colors.Navy;
            row.Cells[0].AddParagraph("REGISTRATION NUMBER");
            row.Cells[0].Format.Font.Bold = false;
            row.Cells[1].AddParagraph("REGISTERATION STATE");
            row.Cells[1].Format.Alignment = ParagraphAlignment.Right;
            row = _table.AddRow();
            row.HeadingFormat = true;
            row.Format.Alignment = ParagraphAlignment.Left;
            row.Format.Font.Bold = true;
            row.Shading.Color = TableBlue;
            row.Cells[0].AddParagraph("EBR52M");
            row.Cells[0].Format.Font.Bold = false;
            row.Cells[1].AddParagraph("NSW");
            row.Cells[1].Format.Alignment = ParagraphAlignment.Right;

            row = _table.AddRow();
            row.HeadingFormat = true;
            row.Format.Alignment = ParagraphAlignment.Left;
            row.Format.Font.Bold = true;
            row.Shading.Color = TableBlue;
            row.Format.SpaceBefore = "0.8cm";
            row.Format.Font.Size = 13;
            row.Format.Font.Color = Colors.Navy;
            row.Cells[0].AddParagraph("REGISTRATION EXPIRY");
            row.Cells[0].Format.Font.Bold = false;
            row = _table.AddRow();
            row.HeadingFormat = true;
            row.Format.Alignment = ParagraphAlignment.Left;
            row.Format.Font.Bold = true;
            row.Shading.Color = TableBlue;
            row.Cells[0].AddParagraph("12/06/2020");
            row.Cells[0].Format.Font.Bold = false;



            row = _table.AddRow();
            row.HeadingFormat = true;
            row.Format.Alignment = ParagraphAlignment.Left;
            row.Format.Font.Bold = true;
            row.Shading.Color = TableBlue;
            row.Format.SpaceBefore = "0.8cm";
            row.Format.Font.Size = 13;
            row.Format.Font.Color = Colors.Navy;
            row.Cells[0].AddParagraph("YEAR OF  MANUFACTURER");
            row.Cells[0].Format.Font.Bold = false;
            row.Cells[1].AddParagraph("YEAR/MONTH OF COMPLIANCE");
            row.Cells[1].Format.Alignment = ParagraphAlignment.Right;
            row = _table.AddRow();
            row.HeadingFormat = true;
            row.Format.Alignment = ParagraphAlignment.Left;
            row.Format.Font.Bold = true;
            row.Shading.Color = TableBlue;
            row.Cells[0].AddParagraph("2001");
            row.Cells[0].Format.Font.Bold = false;
            row.Cells[1].AddParagraph("March 2021");
            row.Cells[1].Format.Alignment = ParagraphAlignment.Right;

            paragraph = section.AddParagraph();
            paragraph.AddText("");
            paragraph.Format.Font.Size = 9;
            paragraph.Format.Alignment = ParagraphAlignment.Center;

            _table2 = section.AddTable();
            _table2.Style = "Table";
            _table2.Borders.Color = Colors.Transparent;
            _table2.Borders.Width = 0.25;
            _table2.Borders.Left.Width = 0.5;
            _table2.Borders.Right.Width = 0.5;
            _table2.Rows.LeftIndent = 0;
            _table2.Format.SpaceBefore = "3.0cm";
            

            column = _table2.AddColumn("16.0cm");
            column.Format.Alignment = ParagraphAlignment.Center;

            row = _table2.AddRow();
            row.Format.SpaceBefore = "3.0cm";
            row.HeadingFormat = true;
            row.Format.Alignment = ParagraphAlignment.Left;
            row.Format.Font.Bold = true;

            row.Format.SpaceBefore = "1.0cm";
            row.Format.Font.Size = 13;
            row.Format.Font.Color = Colors.Navy;
            row.Cells[0].Shading.Color = TableBlue;
            Paragraph para = row.Cells[0].AddParagraph("PPSR REGISTERATION DETAILS");
            para.Format.Alignment = ParagraphAlignment.Center;
            para.Format.LineSpacingRule = LineSpacingRule.Double;
            para.Format.SpaceAfter = 0;
            para.Format.SpaceBefore = 0;
            para.Format.LeftIndent = 0;
            para.Format.FirstLineIndent = "0.5cm";

            row = _table2.AddRow();
            row.Format.SpaceBefore = "3.0cm";
            row.HeadingFormat = true;
            row.Format.Alignment = ParagraphAlignment.Center;
            row.Format.Font.Bold = true;
            para = row.Cells[0].AddParagraph();
            Image image = para.AddImage("greentick.png");
            image.Height = "1.0cm";
            image.Width = "1.0cm";
            Paragraph para2 = row.Cells[0].AddParagraph("There is no security interest or other registration kind registered on the PPSR against the serial number in thesearch criteria details");
            para.Format.LeftIndent = 0;
            para.Format.SpaceBefore = 0;
            para2.Format.LeftIndent = 0;
            para2.Format.SpaceBefore = 0;
            para2.Format.SpaceAfter = "1.0cm";

            row = _table2.AddRow();
            row.Format.SpaceBefore = "3.0cm";
            row.HeadingFormat = true;
            row.Format.Alignment = ParagraphAlignment.Left;
            row.Format.Font.Bold = true;

            row.Format.SpaceBefore = "1.0cm";
            row.Format.Font.Size = 13;
            row.Format.Font.Color = Colors.Navy;
            row.Cells[0].Shading.Color = TableBlue;
            para = row.Cells[0].AddParagraph("NEVDIS WRITTEN OFF VEHICLE NOTIFICATION");
            para.Format.Alignment = ParagraphAlignment.Center;
            para.Format.LineSpacingRule = LineSpacingRule.Double;
            para.Format.SpaceAfter = 0;
            para.Format.SpaceBefore = 0;
            para.Format.LeftIndent = 0;
            para.Format.FirstLineIndent = "0.5cm";
            
            row = _table2.AddRow();
            row.Format.SpaceBefore = "3.0cm";
            row.HeadingFormat = true;
            row.Format.Alignment = ParagraphAlignment.Center;
            row.Format.Font.Bold = true;
            para = row.Cells[0].AddParagraph();
            image = para.AddImage("greentick.png");
            image.Height = "1.0cm";
            image.Width = "1.0cm";
            para2 = row.Cells[0].AddParagraph("Not recorded as written off");
            para.Format.LeftIndent = 0;
            para.Format.SpaceBefore = 0;
            para.Format.SpaceAfter = 0;
            para2.Format.LeftIndent = 0;
            para2.Format.SpaceBefore = 0;
            para2.Format.SpaceAfter = "1.0cm";

            row = _table2.AddRow();
            row.HeadingFormat = true;
            row.Format.Alignment = ParagraphAlignment.Left;
            row.Format.Font.Bold = true;

            row.Format.SpaceBefore = "1.0cm";
            row.Format.Font.Size = 13;
            row.Format.Font.Color = Colors.Navy;
            row.Cells[0].Shading.Color = TableBlue;
            para = row.Cells[0].AddParagraph("NEVDIS STOLEN VEHICLE NOTIFICATION");
            para.Format.Alignment = ParagraphAlignment.Center;

            para.Format.LineSpacingRule = LineSpacingRule.Double;
            para.Format.SpaceAfter = 0;
            para.Format.SpaceBefore = 0;
            para.Format.LeftIndent = 0;
            para.Format.FirstLineIndent = "0.5cm";
            row = _table2.AddRow();
            row.HeadingFormat = true;
            row.Format.Alignment = ParagraphAlignment.Center;
            row.Format.Font.Bold = true;
            para = row.Cells[0].AddParagraph();
            image = para.AddImage("redcross.png");
            image.Height = "1.0cm";
            image.Width = "1.0cm";
            para2 = row.Cells[0].AddParagraph("Its recorded as stolen.  ");
            para2.AddText("Tasmanian stolen vehicle information is currently unavailable from this service. For information about the statusof Tasmanian vehicles, go to:https://www.transport.tas.gov.au/MRSWebInterface/public/regoLookup/registrationLookup.jsf(Note: you willneed to provide the vehicle registration plate number to search this site).A stolen vehicle notification, or the absence of one, does not necessarily mean" +
                " a vehicle is or is not stolen.");
            para.Format.LeftIndent = 0;
            para.Format.SpaceBefore = 3;
            para.Format.SpaceAfter = 0;
            para2.Format.LeftIndent = 0;
            para2.Format.SpaceBefore = 0;
            para2.Format.SpaceAfter = 0;

        }

        /// <summary>
        /// Creates the dynamic parts of the invoice.
        /// </summary>
        void FillContent()
        {
            const double vat = 0.07;

            var paragraph = _addressFrame.AddParagraph();
            paragraph.AddText("Issued to: Shazad Saleemi");
            paragraph.AddLineBreak();
            paragraph.AddText("Email: shazad.saleemi@outlook.com");
            paragraph.AddLineBreak();
            paragraph.AddText("Issued Date: 03 March, 2020");
            paragraph.Format.Alignment = ParagraphAlignment.Right;

        }

        /// <summary>
        /// Selects a subtree in the XML data.
        /// </summary>
        XPathNavigator SelectItem(string path)
        {
            var iter = _navigator.Select(path);
            iter.MoveNext();
            return iter.Current;
        }

        /// <summary>
        /// Gets an element value from the XML data.
        /// </summary>
        static string GetValue(XPathNavigator nav, string name)
        {
            //nav = nav.Clone();
            var iter = nav.Select(name);
            iter.MoveNext();
            return iter.Current.Value;
        }

        /// <summary>
        /// Gets an element value as double from the XML data.
        /// </summary>
        static double GetValueAsDouble(XPathNavigator nav, string name)
        {
            try
            {
                var value = GetValue(nav, name);
                if (value.Length == 0)
                    return 0;
                return Double.Parse(value, CultureInfo.InvariantCulture);
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.Message);
            }
            return 0;
        }

        // Some pre-defined colors...
#if true
        // ... in RGB.
        readonly static Color TableBorder = new Color(81, 125, 192);
        readonly static Color TableBlue = new Color(235, 240, 249);
        readonly static Color TableGray = new Color(242, 242, 242);
#else
        // ... in CMYK.
        readonly static Color TableBorder = Color.FromCmyk(100, 50, 0, 30);
        readonly static Color TableBlue = Color.FromCmyk(0, 80, 50, 30);
        readonly static Color TableGray = Color.FromCmyk(30, 0, 0, 0, 100);
#endif
    }
}

