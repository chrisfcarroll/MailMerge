using System.IO;
using System.Xml;
using System.Xml.XPath;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;

namespace MailMerge.OoXml
{
    public static class StreamAsOoXmlExtensions
    {
        public static 
            WordprocessingDocument 
                AsWordprocessingDocument(this Stream source, bool isEditable=false)
        {
            return WordprocessingDocument.Open(source, isEditable);
        }

        public static string MainDocumentAsXmlString(this WordprocessingDocument document)
        {
            return document.MainDocumentPart.Document.OuterXml;
        }
        
        public static 
            XmlReader 
                MainDocumentAsXmlReader(this WordprocessingDocument document,
                                  FileMode fileMode=FileMode.Open,
                                  FileAccess fileAccess=FileAccess.Read
                                  )
        {
            return XmlReader.Create(document.MainDocumentPart.GetStream(fileMode,fileAccess));
        }

        public static 
            XmlReader 
                AsXmlReaderOfWordprocessingMainDocument(this Stream source, bool isEditable=false)
        {
            var fileAccess =  isEditable?FileAccess.ReadWrite: FileAccess.Read;
            return XmlReader.Create(source.AsWordprocessingDocument().MainDocumentPart.GetStream(FileMode.Open,fileAccess));
        }
        
        public static 
            XPathDocument 
                AsXPathDocOfWordprocessingMainDocument(this Stream source, bool isEditable=false)
        {
            var fileAccess =  isEditable?FileAccess.ReadWrite: FileAccess.Read;
            return new XPathDocument(source.AsWordprocessingDocument().MainDocumentPart.GetStream(FileMode.Open,fileAccess));
        }
    }
}
