using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using iTextSharp;
using iTextSharp.text;
using iTextSharp.text.pdf;
using iTextSharp.text.xml;
using System.IO;
using System.Windows.Documents;

namespace ODTCONVERT
{
    class PDFWriter
    {
        public PDFWriter()
        {
            //Initialization
            DEFAULT_FONTSIZE = 12.0f;
            DEFAULT_HEADINGSIZE = 16.0f;
            FONTNAME = "segoeui.ttf";

            string fg = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Fonts), FONTNAME);
            baseFont = BaseFont.CreateFont(fg, BaseFont.IDENTITY_H, BaseFont.EMBEDDED);
        }

        public void Write(string outputPath, FlowDocument doc)
        {
            Document pdfDoc = new Document(PageSize.A4, 50, 50, 50, 50);
            Font textFont = new Font(baseFont, DEFAULT_FONTSIZE, Font.NORMAL, BaseColor.BLACK);
            Font headingFont = new Font(baseFont, DEFAULT_HEADINGSIZE, Font.BOLD, BaseColor.BLACK);
            using (FileStream stream = new FileStream(outputPath, FileMode.Create))
            {
                PdfWriter writer = PdfWriter.GetInstance(pdfDoc, stream);
                pdfDoc.Open();
                foreach (Block i in doc.Blocks) {
                    if (i is System.Windows.Documents.Paragraph)
                    {
                        TextRange range = new TextRange(i.ContentStart, i.ContentEnd);
                        Console.WriteLine(i.Tag);
                        switch (i.Tag as string)
                        {
                            case "Paragraph": iTextSharp.text.Paragraph par = new iTextSharp.text.Paragraph(range.Text, textFont);
                                par.Alignment = Element.ALIGN_JUSTIFIED;
                                              pdfDoc.Add(par);
                                              break;
                            case "Heading": iTextSharp.text.Paragraph head = new iTextSharp.text.Paragraph(range.Text, headingFont);
                                              head.Alignment = Element.ALIGN_CENTER;
                                              head.SpacingAfter = 10;
                                              pdfDoc.Add(head);
                                              break;
                            default:          iTextSharp.text.Paragraph def = new iTextSharp.text.Paragraph(range.Text, textFont);
                                              def.Alignment = Element.ALIGN_JUSTIFIED;
                                              pdfDoc.Add(def);
                                              break;

                        }
                    }
                    else if (i is System.Windows.Documents.List)
                    {
                        iTextSharp.text.List list = new iTextSharp.text.List(iTextSharp.text.List.UNORDERED, 10f);
                        list.SetListSymbol("\u2022");
                        list.IndentationLeft = 15f;
                        foreach (var li in (i as System.Windows.Documents.List).ListItems)
                        {
                            iTextSharp.text.ListItem listitem = new iTextSharp.text.ListItem();
                            TextRange range = new TextRange(li.Blocks.ElementAt(0).ContentStart, li.Blocks.ElementAt(0).ContentEnd);
                            string text = range.Text.Substring(1);
                            iTextSharp.text.Paragraph par = new iTextSharp.text.Paragraph(text, textFont);
                            listitem.SpacingAfter = 10;
                            listitem.Add(par);
                            list.Add(listitem);
                        }
                        pdfDoc.Add(list);
                    }
                           
               }
                if (pdfDoc.PageNumber == 0)
                {
                    iTextSharp.text.Paragraph par = new iTextSharp.text.Paragraph(" ");
                    par.Alignment = Element.ALIGN_JUSTIFIED;
                    pdfDoc.Add(par);
                }
               pdfDoc.Close();
            }
        }
        BaseFont baseFont;

        public float DEFAULT_FONTSIZE { get; set; }
        public string FONTNAME { get; set; }

        public float DEFAULT_HEADINGSIZE { get; set; }
    }
}
