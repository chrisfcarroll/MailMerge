using System.IO;
using System.Xml;
using System.Xml.Linq;
using System.Xml.XPath;
using DocumentFormat.OpenXml.Packaging;

namespace MailMerge.Helpers
{
    public static class StreamAsOoXmlExtensions
    {
        public static WordprocessingDocument AsWordprocessingDocument(this Stream source, bool isEditable=false)
        {
            source.Position = 0;
            return WordprocessingDocument.Open(source, isEditable);
        }

        public static XPathDocument AsXPathDocOfWordprocessingMainDocument(this Stream source, bool isEditable=false)
        {
            var fileAccess =  isEditable?FileAccess.ReadWrite: FileAccess.Read;
            var docAsXPath = new XPathDocument(source.AsWordprocessingDocument().MainDocumentPart.GetStream(FileMode.Open,fileAccess));
            return docAsXPath;
        }

        /// <summary> XPath query example using LinqToXml / XElement:
        /// <code>using (WordprocessingDocument wordDocument = WordprocessingDocument.Open(filename, true, openSettings))
        /// {
        ///   var xe= XElement.Load(wordDocument.MainDocumentPart.GetStream());
        ///   xe.XPathSelectElements("//w:instrText", OoXmlNamespaces.Manager);
        ///   ... etc ...
        /// 
        /// wordDocument.SaveAs(  Path.Combine(new FileInfo(filename).DirectoryName, new FileInfo(filename).Name + " Output.docx" ) );
        /// }</code>
        /// 
        /// </summary>
        /// <param name="source"></param>
        /// <param name="isEditable"></param>
        /// <returns></returns>
        public static XElement AsXElementOfWordprocessingMainDocument(this Stream source, bool isEditable=false)
        {
            var fileAccess =  isEditable?FileAccess.ReadWrite: FileAccess.Read;
            return XElement.Load(source.AsWordprocessingDocument().MainDocumentPart.GetStream(FileMode.Open,fileAccess));
        }
    }
}
