using System;
using Ionic.Zip;
using System.IO;
using System.Xml;
using System.Windows.Documents;
using System.Windows;

namespace ODTCONVERT
{
    sealed class ODTReader
    {
        // Namespaces. We need this to initialize XmlNamespaceManager so that we can search XmlDocument.
        private static string[,] namespaces = new string[,] 
        {
            {"table", "urn:oasis:names:tc:opendocument:xmlns:table:1.0"},
            {"office", "urn:oasis:names:tc:opendocument:xmlns:office:1.0"},
            {"style", "urn:oasis:names:tc:opendocument:xmlns:style:1.0"},
            {"text", "urn:oasis:names:tc:opendocument:xmlns:text:1.0"},            
            {"draw", "urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"},
            {"fo", "urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0"},
            {"dc", "http://purl.org/dc/elements/1.1/"},
            {"meta", "urn:oasis:names:tc:opendocument:xmlns:meta:1.0"},
            {"number", "urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0"},
            {"presentation", "urn:oasis:names:tc:opendocument:xmlns:presentation:1.0"},
            {"svg", "urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0"},
            {"chart", "urn:oasis:names:tc:opendocument:xmlns:chart:1.0"},
            {"dr3d", "urn:oasis:names:tc:opendocument:xmlns:dr3d:1.0"},
            {"math", "http://www.w3.org/1998/Math/MathML"},
            {"form", "urn:oasis:names:tc:opendocument:xmlns:form:1.0"},
            {"script", "urn:oasis:names:tc:opendocument:xmlns:script:1.0"},
            {"ooo", "http://openoffice.org/2004/office"},
            {"ooow", "http://openoffice.org/2004/writer"},
            {"oooc", "http://openoffice.org/2004/calc"},
            {"dom", "http://www.w3.org/2001/xml-events"},
            {"xforms", "http://www.w3.org/2002/xforms"},
            {"xsd", "http://www.w3.org/2001/XMLSchema"},
            {"xsi", "http://www.w3.org/2001/XMLSchema-instance"},
            {"rpt", "http://openoffice.org/2005/report"},
            {"of", "urn:oasis:names:tc:opendocument:xmlns:of:1.2"},
            {"rdfa", "http://docs.oasis-open.org/opendocument/meta/rdfa#"},
            {"config", "urn:oasis:names:tc:opendocument:xmlns:config:1.0"}
        };

        // Read zip stream (.ods file is zip file).
        private ZipFile GetZipFile(Stream stream)
        {
            return ZipFile.Read(stream);
        }

        // Read zip file (.ods file is zip file).
        private ZipFile GetZipFile(string inputFilePath)
        {
            return ZipFile.Read(inputFilePath);
        }

        private XmlDocument GetContentXmlFile(ZipFile zipFile)
        {
            // Get file(in zip archive) that contains data ("content.xml").
            ZipEntry contentZipEntry = zipFile["content.xml"];

            // Extract that file to MemoryStream.
            Stream contentStream = new MemoryStream();
            contentZipEntry.Extract(contentStream);
            contentStream.Seek(0, SeekOrigin.Begin);
            // Create XmlDocument from MemoryStream (MemoryStream contains content.xml).
            XmlDocument contentXml = new XmlDocument();
            contentXml.Load(contentStream);

            return contentXml;
        }

        private XmlNamespaceManager InitializeXmlNamespaceManager(XmlDocument xmlDocument)
        {
            XmlNamespaceManager nmsManager = new XmlNamespaceManager(xmlDocument.NameTable);

            for (int i = 0; i < namespaces.GetLength(0); i++)
                nmsManager.AddNamespace(namespaces[i, 0], namespaces[i, 1]);

            return nmsManager;
        }
       
        //private XmlNodeList GetTextNodes(XmlDocument contentXmlDocument, XmlNamespaceManager nmsManager)
        //{
        //    return contentXmlDocument.SelectNodes("/office:document-content/office:body/office:text/text:p | /office:document-content/office:body/office:text/text:h", nmsManager);
        //}
        private XmlNodeList GetTextNodes(XmlDocument contentXmlDocument, XmlNamespaceManager nmsManager)
        {
            return contentXmlDocument.SelectNodes("/office:document-content/office:body/office:text/text:*", nmsManager);
        }

        private XmlNodeList GetStyleNodes(XmlDocument contentXmlDocument, XmlNamespaceManager nmsManager)
        {
            return contentXmlDocument.SelectNodes("/office:document-content/office:body/office:text/text:p | /office:document-content/office:body/office:text/text:h", nmsManager);
        }


        private string GetNamespaceUri(string prefix)
        {
            for (int i = 0; i < namespaces.GetLength(0); i++)
            {
                if (namespaces[i, 0] == prefix)
                    return namespaces[i, 1];
            }

            throw new InvalidOperationException("Can't find that namespace URI");
        }

        /// <summary>
        /// Read .odt file and store it in FlowDocument.
        /// </summary>
        /// <param name="inputFilePath">Path to the .odt file.</param>
        /// <returns>FlowDocument that represents .odt file.</returns>
        public FlowDocument Read(string inputFilePath)
        {
            FlowDocument odtFile = new FlowDocument();
            try
            {
                ZipFile odsZipFile = this.GetZipFile(inputFilePath);

                // Get content.xml file
                XmlDocument contentXml = this.GetContentXmlFile(odsZipFile);

                // Initialize XmlNamespaceManager
                XmlNamespaceManager nmsManager = this.InitializeXmlNamespaceManager(contentXml);            

                odtFile.LineStackingStrategy = System.Windows.LineStackingStrategy.BlockLineHeight;

                foreach (XmlNode tableNode in this.GetTextNodes(contentXml, nmsManager))
                {
                    if (tableNode.Name == "text:h")
                    {
                        AddHeading(odtFile, tableNode);
                    }
                    else if (tableNode.Name == "text:p")
                    {
                        AddParagraph(odtFile, tableNode);
                    }
                    else if (tableNode.Name == "text:list")
                    {
                        AddList(odtFile, tableNode);
                    }
                }
            }
            catch (FileNotFoundException e)
            {
                throw;
            }
            catch (Exception e)
            {
                throw new FileFormatException();
            }
            return odtFile;
        }

        private void AddHeading(FlowDocument odtFile, XmlNode tableNode)
        {
            System.Windows.Documents.Paragraph par = new System.Windows.Documents.Paragraph();
            par.Tag = "Heading";
            //Line Spacing
            par.Margin = new Thickness(2);
            //Headings
            par.FontSize = DEFAULT_HEADINGSIZE - Convert.ToInt32(tableNode.Attributes["text:outline-level"].Value) * 2;
            par.TextAlignment = TextAlignment.Center;
            par.Inlines.Add(new Bold(new Run(tableNode.InnerText)));
            odtFile.Blocks.Add(par);
        }

        private void AddParagraph(FlowDocument odtFile, XmlNode tableNode)
        {
            System.Windows.Documents.Paragraph par = new System.Windows.Documents.Paragraph();
            par.Tag = "Paragraph";
            //Line Spacing
            par.Margin = new Thickness(2);
            //Paragraphes
            par.FontSize = DEFAULT_FONTSIZE;
            par.Inlines.Add(new Run(tableNode.InnerText));
            odtFile.Blocks.Add(par);
        }

        private void AddList(FlowDocument odtFile, XmlNode tableNode)
        {
            System.Windows.Documents.List list = new System.Windows.Documents.List();
            foreach (XmlNode listitem in tableNode.ChildNodes)
            {
                if (listitem.Name == "text:list-item")
                {
                    System.Windows.Documents.Paragraph par = new System.Windows.Documents.Paragraph();
                    par.Tag = "ListItem";
                    //Line Spacing
                    par.Margin = new Thickness(2);
                    par.FontSize = DEFAULT_FONTSIZE;
                    par.Inlines.Add(new Run(listitem.InnerText));
                    System.Windows.Documents.ListItem item = new System.Windows.Documents.ListItem(par);
                    list.ListItems.Add(item);
                }
            }
            odtFile.Blocks.Add(list);
        }
        public ODTReader(){
            //Initialization
            DEFAULT_FONTSIZE = 16.0;
            DEFAULT_HEADINGSIZE = 30.0;
        }

        public double DEFAULT_FONTSIZE { get; set; }

        public double DEFAULT_HEADINGSIZE { get; set; }
    }
}
