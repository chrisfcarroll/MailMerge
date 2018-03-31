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
        public static void SimpleMergeFields(this XmlDocument mainDocumentPart, Dictionary<string, string> fieldValues, ILogger logger)
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
        }

        public static void ComplexMergeFields(this XmlDocument mainDocumentPart, Dictionary<string, string> fieldValues, ILogger logger)
        {
            int expectedNodeCount = 5;

            XmlNode beginRun;
            while (null!= (beginRun= mainDocumentPart.SelectSingleNode("//w:r[w:fldChar/@w:fldCharType='begin']", OoXmlNamespaces.Manager)))
            {
                var nodesToRemove= new List<XmlNode>{beginRun};
                XmlNode instrRun=null, instrNode=null, separatorRun = null, textRun = null, endRun = null;
                string replacementText = "";

                int i = 0;
                var sibling = beginRun.NextSibling;
                while (endRun== null && i<expectedNodeCount && sibling!=null)
                {
                    if (null != sibling.SelectSingleNode("w:fldChar[@w:fldCharType='separate']", OoXmlNamespaces.Manager))
                    {
                        nodesToRemove.Add(separatorRun = sibling);
                    }
                    if (null != sibling.SelectSingleNode("w:fldChar[@w:fldCharType='end']", OoXmlNamespaces.Manager))
                    {
                        nodesToRemove.Add(endRun=sibling);
                    }
                    if (null != (instrNode=sibling.SelectSingleNode("w:instrText[contains(text(),'MERGEFIELD ')]", OoXmlNamespaces.Manager)))
                    {
                        nodesToRemove.Add(instrRun = sibling);
                        var fieldName = instrNode.InnerText
                                                .Split(" ", StringSplitOptions.RemoveEmptyEntries)
                                                .Skip(1).FirstOrDefault();
                        if (fieldValues.ContainsKey(fieldName))
                        {
                            replacementText = fieldValues[fieldName];
                            logger.LogDebug($"Noting <w:instrText '{fieldName} '>...</w:instrText> for replacement with " + replacementText);
                        }
                        else
                        {
                            logger.LogWarning($"Nothing in the fieldValue dictionary for {instrRun.InnerText}");
                        }

                    }
                    if (endRun==null && separatorRun!=null /*17.16.18 only do replacement after a separator. no separator implies no replacement*/)
                    {
                        textRun = sibling;
                    }
                    sibling = sibling.NextSibling;
                    i++;
                }

                if (textRun != null )
                {
                    textRun.SelectSingleNode("w:t",OoXmlNamespaces.Manager).InnerText = replacementText;
                }
                nodesToRemove.ForEach(n=>n.ParentNode.RemoveChild(n));
            }
        }

        /// <summary>
        /// ECMA-376 Part 1 
        /// 17.16 Fields and Hyperlinks
        /// TODO : 17.16.4.1 Date and time formatting only partially implemented. Where the standard overlaps with .Net standard, it will work.
        /// </summary>
        /// <param name="mainDocumentPart">The document to mutate</param>
        /// <param name="formattedFixedDate">A literal string that will be used to replace date fields</param>
        /// <param name="logger"></param>
        /// <param name="datesToReplace">Defaults to <seealso cref="DefaultDatesToReplace"/>, i.e. <code>{"DATE", "PRINTDATE", "SAVEDATE"}</code>
        /// Possible item values are date-and-time= "CREATEDATE" | "DATE" | "EDITTIME" | "PRINTDATE" | "SAVEDATE" | "TIME"</param>
        public static void MergeDate(this XmlDocument mainDocumentPart, ILogger logger, string formattedFixedDate, string[] datesToReplace=default)
        {
            MergeDate(mainDocumentPart, logger, DateTime.Now, formattedFixedDate, datesToReplace);
        }

        /// <summary>
        /// ECMA-376 Part 1 
        /// 17.16 Fields and Hyperlinks
        /// TODO : 17.16.4.1 Date and time formatting only partially implemented. Where the standard overlaps with .Net standard, it will work.
        /// </summary>
        /// <param name="mainDocumentPart">The document to mutate</param>
        /// <param name="logger"></param>
        /// <param name="date">The date which will be used to replace date fields. If the date field formatting instruction can be parsed, 
        /// it will be used. If not, then <seealso cref="DateTime.ToLongDateString"/> will be applied to format the replacement date.
        /// </param>
        /// <param name="formattedFixedDate">A literal string that will be used to replace date fields. 
        /// If this is not null, it overrides both <paramref name="date"/> and any formatting instructions in the date merge fields
        /// </param>
        /// <param name="datesToReplace">Defaults to <seealso cref="DefaultDatesToReplace"/>, i.e. <code>{"DATE", "PRINTDATE", "SAVEDATE"}</code>
        /// Possible item values are date-and-time= "CREATEDATE" | "DATE" | "EDITTIME" | "PRINTDATE" | "SAVEDATE" | "TIME"
        /// </param>
        public static void MergeDate(this XmlDocument mainDocumentPart, ILogger logger, DateTime? date, string formattedFixedDate = null, string[] datesToReplace=null)
        {
            datesToReplace = datesToReplace ?? DefaultDatesToReplace;
            foreach (var dateType in datesToReplace)
            {
                XmlNode node;
                while (null != (node = mainDocumentPart.SelectSingleNode($"//w:instrText[contains(text(),'{dateType} ')]", OoXmlNamespaces.Manager)))
                {
                    string format;
                    if (formattedFixedDate==null )
                    {
                        var left = node.InnerText.IndexOf("\\@") + 2;
                        if (left < 2) left = 0;
                        var right= node.InnerText.IndexOf('\\',left);
                        if (right < 0) right= node.InnerText.Length;
                        format = node.InnerText.Substring(left, right - left).Trim(' ','"','\\');
                        if(format.Length>0)
                        try
                        {
                            //TODO: WordProcessingML date formats are not the same as C# formats. This is a limited translation
                            format = format.Replace('D', 'd').Replace('b', 'y').Replace('B', 'y').Replace('A', 'd');
                            formattedFixedDate = (date??DateTime.Now).ToString(format);
                        }
                        catch{}
                    }

                    logger.LogDebug($"Replacing <w:instrText '{dateType} '>...</w:instrText> with " + formattedFixedDate);
                    var replacementNode = mainDocumentPart.CreateElement("w", "t", OoXmlNamespaces.Instance["w"]);
                    replacementNode.InnerText = formattedFixedDate ?? (date??DateTime.Now).ToLongDateString();
                    node.ParentNode.ReplaceChild(replacementNode, node);
                }
            }
        }

        public static readonly string[] DefaultDatesToReplace = {"DATE", "PRINTDATE", "SAVEDATE"};
    }
}
