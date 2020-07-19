using System.IO;
using System.Xml;
using System.Xml.Linq;

namespace MailMerge.Helpers
{
    /// <summary>
    /// Extension methods to read filesystem files as OOXml WordProcessing Documents
    /// </summary>
    public static class FileInfoAsOoxmlExtensions
    {
        /// <summary>
        /// Read the OoXml Wordprocessing <see cref="DocumentFormat.OpenXml.Packaging.MainDocumentPart"/>
        /// of <paramref name="file"/> as an <see cref="XElement"/>
        /// </summary>
        /// <param name="file"></param>
        /// <returns>A <see cref="XElement"/> containing the MainDocumentPart</returns>
        /// <remarks><paramref name="file"/> is opened with Read Access only.</remarks>
        public static XElement GetXElementOfWordprocessingMainDocument(this FileInfo file)
        {
            using var fileStream= file.OpenRead();
            return fileStream.AsXElementOfWordprocessingMainDocument(isEditable:false);
        }
        
        /// <summary>
        /// Read the OoXml Wordprocessing <see cref="DocumentFormat.OpenXml.Packaging.MainDocumentPart"/>
        /// of <paramref name="file"/> as an <see cref="XmlDocument"/> 
        /// </summary>
        /// <param name="file"></param>
        /// <returns>A <see cref="XmlDocument"/> containing the MainDocumentPart</returns>
        /// <remarks><paramref name="file"/> is opened with Read Access only.</remarks>
        public static XmlDocument GetXmlDocumentOfWordprocessingMainDocument(this FileInfo file)
        {
            using var fileStream= file.OpenRead();
            return fileStream.GetXmlDocumentOfWordprocessingMainDocument();
        }
    }
}