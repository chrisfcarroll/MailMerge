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
        /// ECMA-376 Part 1 17.16.5.35
        /// <![CDATA[<w:fldSimple w:instr=" MERGEFIELD Name "><w:t>«Name»</w:t></w:fldSimple>]]>  
        /// </summary>
        /// <param name="mainDocumentPart">The document to mutate</param>
        /// <param name="fieldValues">The dictionary of MERGEFIELD values to use for replacement</param>
        /// <param name="log"></param>
        /// <example><![CDATA[
        /// <w:fldSimple w:instr=" MERGEFIELD FirstName " xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
        /// 	<w:r w:rsidR="00F73BE2">
        /// 		<w:rPr><w:noProof /></w:rPr>
        /// 		<w:t>«FirstName»</w:t>
        /// 	</w:r>
        /// </w:fldSimple>]]></example>
        public static void SimpleMergeFields(this XmlDocument mainDocumentPart, Dictionary<string, string> fieldValues, ILogger log)
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
                        log.LogDebug($"Replacing <w:fldSimple w:instr='MERGEFIELD {fieldName}'>...<w:t>{txtNode.InnerText}</w:t> with {fieldValues[fieldName]}");
                        txtNode.InnerText = fieldValues[fieldName];
                    }
                }
            }
        }

        /// <summary>ECMA-376 Part 1 17.16  <![CDATA[<w:instrText> MERGEFIELD</w:instrText>]]>
        /// Each <![CDATA[<w:instrText> MERGEFIELD</w:instrText>]]> is found inside a sequence of 5 or more <![CDATA[<w:r >]]>
        /// nodes, known as Runs:
        /// <list type="bullet">
        /// <item><![CDATA[<w:r><w:fldChar w:fldCharType='begin'></w:r>]]> Begin Run</item>
        /// <item><![CDATA[<w:r><w:instrText></w:r>]]> - 1 or more of these which may need to be stuck back together</item>
        /// <item><![CDATA[<w:r><w:fldChar fldCharType='separator'></w:r>]]>Separator Run</item>
        /// <item><![CDATA[<w:r><w:t></w:r>]]>Text Run</item>
        /// <item><![CDATA[<w:r><w:fldChar fldCharType='end'></w:r>]]>End Run</item>
        /// </list>
        /// We update the textRun from fieldValues, and remove the other Runs.
        /// </summary>
        /// <param name="mainDocumentPart"></param>
        /// <param name="fieldValues">Dictionary of merge replacement fields</param>
        /// <param name="log"></param>
        /// <example>
        /// <![CDATA[
        /// <w:r>
        ///    <w:fldChar w:fldCharType="begin"/>
        ///  </w:r>
        /// <!-- The <w:r><w:instrText></w:r> node may be split out over several <w:r> nodes -->
        ///  <w:r>
        ///    <w:instrText xml:space="preserve">MERGEFIELD  Name  \* MERGEFORMAT </w:instrText>
        ///  </w:r>
        ///  <w:r>
        ///    <w:fldChar w:fldCharType="separate"/>
        ///  </w:r>
        ///  <w:r>
        ///    <w:t>«Name»</w:t>
        ///  </w:r>
        ///  <w:r>
        ///    <w:fldChar w:fldCharType="end"/>
        ///  </w:r>]]></example>
        public static void ComplexMergeFields(this XmlDocument mainDocumentPart, Dictionary<string, string> fieldValues, ILogger log)
        {
            const int boilerplateCount = 3; 

            var beginRuns= mainDocumentPart.SelectNodes("//w:r[w:fldChar/@w:fldCharType='begin']", OoXmlNamespaces.Manager);
            log.LogDebug("Found " + beginRuns.Count + " <w:fldChar w:fldCharType='begin'> nodes.");
            foreach(XmlNode beginRun in beginRuns)
            {
                var boilerPlateNodes= new List<XmlNode>{beginRun};
                var instrRuns= new List<XmlNode>();
                XmlNode separatorRun = null, textRun = null, instrNode;
                string replacementText = "";
                bool statePendingFieldName = false;
                int i = 0;

                var sibling = beginRun;
                while (   (sibling = sibling.NextSibling)!= null 
                       && (++i < boilerplateCount + instrRuns.Count ))
                {
                    if (null != sibling.SelectSingleNode("w:fldChar[@w:fldCharType='end']", OoXmlNamespaces.Manager))
                    {
                        boilerPlateNodes.Add(sibling);
                        break;
                    }
                    else if (null != sibling.SelectSingleNode("w:fldChar[@w:fldCharType='separate']", OoXmlNamespaces.Manager))
                    {
                        boilerPlateNodes.Add(separatorRun = sibling);
                    }
                    else if (separatorRun != null /*17.16.18 only replace after a separator; no separator=no replace*/
                        && sibling.SelectNodes("w:t", OoXmlNamespaces.Manager).Count>0 )
                    {
                        textRun = sibling;
                    }
                    else if (null != (instrNode=sibling.SelectSingleNode("w:instrText[contains(text(),'MERGEFIELD ')]", OoXmlNamespaces.Manager)))
                    {
                        instrRuns.Add(sibling);
                        var fieldName = instrNode.InnerText
                                                .Split(" ", StringSplitOptions.RemoveEmptyEntries)
                                                .Skip(1).FirstOrDefault();
                        if (fieldName==null)
                        {
                            statePendingFieldName = true;
                            log.LogDebug("Noting <w:instrText MERGEFIELD *with no FieldName* >...</w:instrText> for potential replacement");
                        }
                        else if (fieldValues.ContainsKey(fieldName))
                        {
                            replacementText = fieldValues[fieldName];
                            log.LogDebug($"Noting <w:instrText '{fieldName} '>...</w:instrText> for replacement with " + replacementText);
                        }
                        else
                        {
                            log.LogWarning($"Nothing in the fieldValue dictionary for {sibling.InnerText}");
                        }
                    }
                    else if (statePendingFieldName && null != (instrNode=sibling.SelectSingleNode("w:instrText[not( contains(text(),'MERGEFIELD '))]", OoXmlNamespaces.Manager)))
                    {
                        statePendingFieldName = false;
                        instrRuns.Add(sibling);
                        var fieldName = instrNode.InnerText
                            .Split(" ", StringSplitOptions.RemoveEmptyEntries)
                            .FirstOrDefault();

                        if (fieldValues.ContainsKey(fieldName))
                        {
                            replacementText = fieldValues[fieldName];
                            log.LogDebug($"Noting <w:instrText '{fieldName} '>...</w:instrText> for replacement with " + replacementText);
                        }
                        else
                        {
                            log.LogWarning($"Nothing in the fieldValue dictionary for {sibling.InnerText}");
                        }
                    }
                }

                if (textRun != null )
                {
                    textRun
                        .SelectSingleNode("w:t",OoXmlNamespaces.Manager).InnerText = replacementText;
                    boilerPlateNodes
                        .ForEach(n=>n.RemoveMe());
                    instrRuns
                        .ForEach(n=> n.RemoveMe());
                }
                else if(instrRuns.Any())
                {
                    log.LogWarning($"Ignored sequence containing {instrRuns.Last()} because it was incomplete");
                }
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
