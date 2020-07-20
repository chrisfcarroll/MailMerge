using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using MailMerge.Helpers;
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
    }
}