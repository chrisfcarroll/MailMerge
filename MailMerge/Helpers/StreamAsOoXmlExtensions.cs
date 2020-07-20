using System;
using System.IO;
using System.Xml;
using System.Xml.Linq;
using System.Xml.XPath;
using DocumentFormat.OpenXml.Packaging;

namespace MailMerge
{
    /// <summary>
    /// Extension Methods to read streams as OoXml documents
    /// </summary>
    public static class StreamAsOoXmlExtensions
    {
        /// <summary>
        /// A convenience method to wrap a <see cref="WordprocessingDocument"/> around an open stream
        /// having first reset the stream position to 0. A shortcut for e.g.
        /// <code>
        /// stream.Position=0;
        /// using (WordprocessingDocument wordDocument = WordprocessingDocument.Open(stream, isEditable, openSettings))
        /// </code>
        /// The returned <see cref="WordprocessingDocument"/> can subsequently be saved:
        /// <code>
        /// wordDocument.SaveAs(  Path.Combine(new FileInfo(filename).DirectoryName, new FileInfo(filename).Name + " Output.docx" ) );
        /// </code>
        /// </summary>
        /// <param name="source"></param>
        /// <param name="isEditable"></param>
        /// <returns></returns>
        /// <remarks><see cref="WordprocessingDocument"/> implements <see cref="IDisposable"/></remarks>
        public static WordprocessingDocument AsWordprocessingDocument(this Stream source, bool isEditable=false)
        {
            source.Position = 0;
            return WordprocessingDocument.Open(source, isEditable);
        }

        /// <summary>
        /// Read the OoXml Wordprocessing <see cref="DocumentFormat.OpenXml.Packaging.MainDocumentPart"/>
        /// of <paramref name="source"/> as an <see cref="XPathDocument"/> 
        /// </summary>
        /// <param name="source">A stream of an
        /// <see cref="DocumentFormat.OpenXml.Packaging.WordprocessingDocument"/></param>
        /// <param name="isEditable">If this is set <c>true</c> then the <paramref name="source"/>
        /// may be edited by operations on the returned <see cref="XElement"/></param>
        /// <returns>An <see cref="XElement"/> containing the MainDocumentPart</returns>
        ///<remarks>
        /// XPath query example using LinqToXml / XElement:
        /// <code>using stream= fileInfo.OpenRead();
        /// var xe= stream.AsXPathDocOfWordprocessingMainDocument();
        /// xe.XPathSelectElements("//w:instrText", OoXmlNamespaces.Manager);
        ///   ... etc ...
        /// </code>
        /// </remarks>
        public static XPathDocument AsXPathDocOfWordprocessingMainDocument(this Stream source, bool isEditable=false)
        {
            var fileAccess =  isEditable?FileAccess.ReadWrite: FileAccess.Read;
            var docAsXPath = new XPathDocument(source.AsWordprocessingDocument().MainDocumentPart.GetStream(FileMode.Open,fileAccess));
            return docAsXPath;
        }

        /// <summary>
        /// Read the OoXml Wordprocessing <see cref="DocumentFormat.OpenXml.Packaging.MainDocumentPart"/>
        /// of <paramref name="source"/> as an <see cref="XmlDocument"/> 
        /// </summary>
        /// <param name="source">A stream of an
        /// <see cref="DocumentFormat.OpenXml.Packaging.WordprocessingDocument"/></param>
        /// <param name="isEditable">If this is set <c>true</c> then the <paramref name="source"/>
        /// may be edited by operations on the returned <see cref="XElement"/></param>
        /// <returns>An <see cref="XElement"/> containing the MainDocumentPart</returns>
        public static XElement AsXElementOfWordprocessingMainDocument(this Stream source, bool isEditable=false)
        {
            var fileAccess =  isEditable?FileAccess.ReadWrite: FileAccess.Read;
            return XElement.Load(source.AsWordprocessingDocument().MainDocumentPart.GetStream(FileMode.Open,fileAccess));
        }

        /// <summary>
        /// Read the OoXml Wordprocessing <see cref="DocumentFormat.OpenXml.Packaging.MainDocumentPart"/>
        /// of <paramref name="source"/> as an <see cref="XmlDocument"/> 
        /// </summary>
        /// <param name="source"></param>
        /// <returns></returns>
        public static XmlDocument GetXmlDocumentOfWordprocessingMainDocument(this Stream source)
        {
            var xdoc = new XmlDocument(OoXmlNamespace.Manager.NameTable);
            using (var wpDocx = WordprocessingDocument.Open(source, false))
            using (var docOutStream = wpDocx.MainDocumentPart.GetStream(FileMode.Open, FileAccess.Read))
            {
                xdoc.Load(docOutStream);
            }
            return xdoc;
        }
    }
}
