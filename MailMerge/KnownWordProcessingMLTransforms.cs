using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml;
using MailMerge.Helpers;
using Microsoft.Extensions.Logging;

namespace MailMerge
{

    public static class KnownWordProcessingMLTransforms
    {
        /// <summary>
        /// ECMA-376 Part 1 17.16.5.35 MERGEFIELD  
        /// </summary>
        /// <param name="mainDocumentPart">The document to mutate</param>
        /// <param name="fieldValues">The dictionary of MERGEFIELD values to use for replacement</param>
        /// <param name="logger"></param>
        public static void MergeField(this XmlDocument mainDocumentPart, Dictionary<string, string> fieldValues, ILogger logger)
        {
            var simpleMergeFields = mainDocumentPart.SelectNodes("//w:fldSimple[contains(@w:instr,'MERGEFIELD ')]", OoXmlNamespaces.Manager);
            foreach (XmlNode node in simpleMergeFields)
            {
                var fieldName = node.Attributes[OoXPath.winstr].Value
                                        .Split(" ", StringSplitOptions.RemoveEmptyEntries)
                                        .Skip(1).FirstOrDefault();
                if (fieldValues.ContainsKey(fieldName))
                {
                    foreach (XmlNode txtNode in node.SelectNodes(".//w:t", OoXmlNamespaces.Manager))
                    {
                        logger.LogDebug($"Replacing <w:fldSimple w:instr='MERGEFIELD {fieldName}'>...<w:t>{txtNode.InnerText}</w:t> with {fieldValues[fieldName]}");
                        txtNode.InnerText = fieldValues[fieldName];
                    }
                }
            }

            //While loop not foreach because of deleting as we go.
            var winstrNode = mainDocumentPart.SelectSingleNode("//w:instrText[contains(text(),'MERGEFIELD ')]", OoXmlNamespaces.Manager);
            while (winstrNode != null)
            {
                var fieldName = winstrNode.InnerText
                    .Split(" ", StringSplitOptions.RemoveEmptyEntries)
                    .Skip(1).FirstOrDefault();
                if (fieldValues.ContainsKey(fieldName))
                {
                    logger.LogDebug($"Replacing <w:instrText '{fieldName} '>...</w:instrText> with " + fieldValues[fieldName]);
                    winstrNode.ParentNode.InsertAfter(mainDocumentPart.CreateElement($"<w:t>{fieldValues[fieldName]}</w:t>"), winstrNode);
                    winstrNode.ParentNode.RemoveChild(winstrNode);
                }
                winstrNode = mainDocumentPart.SelectSingleNode("//w:instrText[contains(text(),'MERGEFIELD ')]", OoXmlNamespaces.Manager);
            }
        }

        // 
        /// <summary>
        /// ECMA-376 Part 1 
        /// 17.16 Fields and Hyperlinks
        /// TODO : 17.16.4.1 Date and time formatting only partially implemented. Where the standard overlaps with .Net standard, it will work.
        /// </summary>
        /// <param name="mainDocumentPart">The document to mutate</param>
        /// <param name="date">The date to use. Defaults to <seealso cref="DateTime.Today"/>.<seealso cref="DateTime.ToLongDateString"/>()  </param>
        /// <param name="logger"></param>
        /// <param name="datesToReplace">Defaults to <seealso cref="DefaultDatesToReplace"/>, i.e. <code>{"DATE", "PRINTDATE", "SAVEDATE"}</code>
        /// Possible item values are date-and-time= "CREATEDATE" | "DATE" | "EDITTIME" | "PRINTDATE" | "SAVEDATE" | "TIME"</param>
        public static void MergeDate(this XmlDocument mainDocumentPart, ILogger logger, string date=default, string[] datesToReplace=default)
        {
            datesToReplace = datesToReplace ?? DefaultDatesToReplace;
            date = date == default ? DateTime.Today.ToLongDateString() : date;
            foreach (var dateType in datesToReplace)
            {
                var dateFields = mainDocumentPart.SelectNodes( $"//w:instrText[contains(text(),'{dateType} ')]", OoXmlNamespaces.Manager);
                foreach (XmlNode node in dateFields)
                {
                    var formattedDate = date;
                    var format = node.InnerText.Split('\\', StringSplitOptions.RemoveEmptyEntries).Skip(1).FirstOrDefault()?.Trim(' ','"','\'');
                    if (format != null){
                        try
                        {
                            //TODO: WordProcessingML date formats are NOT the same as C# formats. This is a limited translation
                            format = format.Replace('D', 'd').Replace('b','y').Replace('B','y').Replace('A','d');
                            formattedDate = DateTime.Now.ToString(format);
                        }catch{}}
                    logger.LogDebug($"Replacing <w:instrText '{dateType} '>...</w:instrText> with " + formattedDate);
                    node.InnerText = formattedDate;
                }
            }
        }

        public static readonly string[] DefaultDatesToReplace = {"DATE", "PRINTDATE", "SAVEDATE"};
    }
}
