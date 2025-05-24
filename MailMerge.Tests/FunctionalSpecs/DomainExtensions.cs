using DocumentFormat.OpenXml.Packaging;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using TestBase;

namespace MailMerge.Tests.FunctionalSpecs
{
    static class DomainExtensions
    {
        public static IEnumerable<string> 
            ShouldHaveBeenSplitOnNewLinesAndEachLineInserted(
                this IEnumerable<string> multiLineSourceFields, string outputXml)
        {
            return multiLineSourceFields.Each(s => ShouldHaveBeenSplitOnNewLinesAndEachLineInserted((string) s, outputXml));
        }

        public static string ShouldHaveBeenSplitOnNewLinesAndEachLineInserted(this string field, string outputXml)
        {
            var lines = field.Split(new[] {"\n", "\n\r"}, StringSplitOptions.None);
            outputXml.ShouldContain(lines[0]);
            foreach (var line in lines.Skip(1))
            {
                outputXml.ShouldContain($"<w:br />{line}");
            }

            return field;
        }

        public static (string asText, string asXml) AsWordDocumentMainPartTextAndXml(this Stream documentStream)
        {
            var asText = documentStream.AsWordprocessingDocument().MainDocumentPart.Document.InnerText;
            var asXml = documentStream.AsWordprocessingDocument().MainDocumentPart.Document.OuterXml;
            return (asText, asXml);

        }

        public static (string allHeaderText, string allHeaderXml) GetCombinedHeadersTextAndXml(this Stream documentStream)
        {
            var sbHeaderText = new StringBuilder();
            var sbHeaderXml = new StringBuilder();

            if (documentStream.CanSeek)
            {
                documentStream.Position = 0;
            }

            using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(documentStream, isEditable: false))
            {
                if (wordDoc.MainDocumentPart != null)
                {
                    foreach (HeaderPart headerPart in wordDoc.MainDocumentPart.HeaderParts)
                    {
                        if (headerPart.Header != null)
                        {
                            sbHeaderText.Append(headerPart.Header.InnerText);
                            sbHeaderXml.Append(headerPart.Header.OuterXml);
                        }
                    }
                }
            }
            return (sbHeaderText.ToString(), sbHeaderXml.ToString());
        }
    }
}