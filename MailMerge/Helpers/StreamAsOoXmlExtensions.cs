using System.IO;
using System.Xml.XPath;
using DocumentFormat.OpenXml.Packaging;

namespace MailMerge.Helpers
{
    public static class StreamAsOoXmlExtensions
    {
        public static 
            WordprocessingDocument 
                AsWordprocessingDocument(this Stream source, bool isEditable=false)
        {
            source.Position = 0;
            return WordprocessingDocument.Open(source, isEditable);
        }

        public static 
            XPathDocument 
                AsXPathDocOfWordprocessingMainDocument(this Stream source, bool isEditable=false)
        {
            var fileAccess =  isEditable?FileAccess.ReadWrite: FileAccess.Read;
            var docAsXPath = new XPathDocument(source.AsWordprocessingDocument().MainDocumentPart.GetStream(FileMode.Open,fileAccess));
            return docAsXPath;
        }
    }
}
