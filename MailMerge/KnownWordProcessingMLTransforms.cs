using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml;
using Microsoft.Extensions.Logging;

namespace MailMerge
{

    /// <summary>
    /// <para>KnownWordProcessingMLTransforms - a subset of OfficeOpenXml that should be sufficient for
    /// programmatically merging documents.</para>
    /// <para>For a helpful, explanatory, mini-reference: http://officeopenxml.com/</para>
    /// <para>For the full specification: https://www.ecma-international.org/publications/standards/Ecma-376.htm</para>
    /// </summary>
    /// <remarks>
    /// <para>For a helpful, explanatory, mini-reference: http://officeopenxml.com/</para>
    /// <para>For the full specification: https://www.ecma-international.org/publications/standards/Ecma-376.htm</para>
    /// </remarks>
    public static class KnownWordProcessingMLTransforms
    {
        public static readonly string[] DefaultDatesToReplace = {"DATE", "PRINTDATE", "SAVEDATE"};
        public static readonly string[] NewLineSeparators =  {"\n", "\n\r"};
        
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
        /// <remarks>
        /// ECMA-376 Part 1 17.16.2 XML representation
        /// Fields shall be implemented in XML using either of two approaches:
        /// • As a simple field implementation, using the fldSimple element, or
        /// • As a complex field implementation, using a set of runs involving the fldChar and instrText elements.
        /// </remarks>
        public static void SimpleMergeFields(this XmlDocument mainDocumentPart, Dictionary<string, string> fieldValues, ILogger log)
        {
            var simpleMergeFields = mainDocumentPart.SelectNodes("//w:fldSimple[contains(@w:instr,'MERGEFIELD ')]", OoXmlNamespace.Manager);
            foreach (XmlNode node in simpleMergeFields)
            {
                var fieldName = node.Attributes[OoXPath.winstr].Value
                                        .Split(" ", StringSplitOptions.RemoveEmptyEntries)
                                        .Skip(1).FirstOrDefault();
                if (fieldValues.ContainsKey(fieldName))
                {
                    foreach (XmlNode txtNode in node.SelectNodes(".//w:t", OoXmlNamespace.Manager))
                    {
                        log.LogDebug($"Replacing <w:fldSimple w:instr='MERGEFIELD {fieldName}'>...<w:t>{txtNode.InnerText}</w:t> with {fieldValues[fieldName]}");
                        //txtNode.InnerText = fieldValues[fieldName];
                        txtNode.ReplaceInnerText(fieldValues[fieldName],mainDocumentPart,log);
                    }
                }
            }
        }

        /// <summary>ECMA-376 Part 1 17.16  <![CDATA[<w:instrText> MERGEFIELD</w:instrText>]]>
        /// Each <![CDATA[<w:instrText> MERGEFIELD</w:instrText>]]> is found inside a sequence of 5 or more
        /// runs, <![CDATA[<w:r >]]> nodes:
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
        /// <remarks>
        /// ECMA-376 Part 1 17.16.2 XML representation
        /// Fields shall be implemented in XML using either of two approaches:
        /// • As a simple field implementation, using the fldSimple element, or
        /// • As a complex field implementation, using a set of runs involving the fldChar and instrText elements.
        /// </remarks>

        public static void ComplexMergeFields(this XmlDocument mainDocumentPart, Dictionary<string, string> fieldValues, ILogger log)
        {
            const int boilerplateCount = 3; 

            var beginRuns= mainDocumentPart.SelectNodes("//w:r[w:fldChar/@w:fldCharType='begin']", OoXmlNamespace.Manager);
            log.LogDebug("Found " + beginRuns.Count + " <w:fldChar w:fldCharType='begin'> nodes.");
            foreach(XmlNode beginRun in beginRuns)
            {
                var boilerPlateNodes= new List<XmlNode>{beginRun};
                var instrRuns= new List<XmlNode>();
                XmlNode separatorRun = null, textRun = null, instrNode;
                string replacementText = "";
                int i = 0;

                var sibling = beginRun;
                while (   (sibling = sibling.NextSibling)!= null 
                       && (++i < boilerplateCount + instrRuns.Count ))
                {
                    if (null != sibling.SelectSingleNode("w:fldChar[@w:fldCharType='end']", OoXmlNamespace.Manager))
                    {
                        boilerPlateNodes.Add(sibling);
                        break;
                    }
                    else if (null != sibling.SelectSingleNode("w:fldChar[@w:fldCharType='separate']", OoXmlNamespace.Manager))
                    {
                        boilerPlateNodes.Add(separatorRun = sibling);
                    }
                    else if (separatorRun != null /*17.16.18 only replace after a separator; no separator=no replace*/
                        && sibling.SelectNodes("w:t", OoXmlNamespace.Manager).Count>0 )
                    {
                        textRun = sibling;
                    }
                    else if (null != (instrNode=sibling.SelectSingleNode("w:instrText[contains(text(),'MERGEFIELD ')]", OoXmlNamespace.Manager)))
                    {
                        var nodesText = CollectSequentialInstrRuns(instrNode, ref sibling, instrRuns);

                        var fieldName = nodesText
                                                .Split(" ", StringSplitOptions.RemoveEmptyEntries)
                                                .Skip(1).FirstOrDefault();
                        if (fieldName==null)
                        {
                            log.LogWarning("<w:instrText MERGEFIELD *with no FieldName* >...</w:instrText> cannot be merged");
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
                }

                if (textRun != null )
                {
                    textRun.ReplaceInnerText(replacementText, mainDocumentPart, log);
                    boilerPlateNodes
                        .ForEach(n=>n.RemoveMe());
                    instrRuns
                          .ForEach(n => n.RemoveMe());
                }
                else if(instrRuns.Any())
                {
                    log.LogWarning($"Ignored sequence containing {instrRuns.Last()} because it was incomplete");
                }
            }
        }

        /// <summary>Collects sequential <![CDATA[<w:instrText></w:instrText>]]>
        /// that do not contain the text "MERGEFIELD " and joins the inner text.
        /// <example>
        /// <![CDATA[
        ///  <w:r>
        ///    <w:instrText xml:space="preserve">MERGEFIELD  Name  \* M</w:instrText>
        ///  </w:r>
        ///  <w:r>
        ///    <w:instrText xml:space="preserve">ERGEFORMAT </w:instrText>
        ///  </w:r>
        ///  ]]></example>
        /// <param name="instrNode">first "instrText" node</param>
        /// <param name="sibling">run node containing the first "instrText" node, will be updated to the run containing the last sequential "instText" node</param>
        /// <param name="instrRuns">list of instrText runs, runs will be added to this list</param>
        /// <returns>Concatenation of all sequential instrText nodes InnerText</returns>
        private static string CollectSequentialInstrRuns(XmlNode instrNode, ref XmlNode sibling, List<XmlNode> instrRuns)
        {
            instrRuns.Add(sibling);
            var nodesText = instrNode.InnerText;
            XmlNode instrSibling = sibling.NextSibling;

            while (null != instrSibling && null != (instrNode = instrSibling.SelectSingleNode("w:instrText[not( contains(text(),'MERGEFIELD '))]", OoXmlNamespace.Manager)))
            {
                nodesText += instrNode.InnerText;
                instrRuns.Add(instrSibling);
                sibling = instrSibling;
                instrSibling = instrSibling.NextSibling;
            }

            return nodesText;
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
                while (null != (node = mainDocumentPart.SelectSingleNode($"//w:instrText[contains(text(),'{dateType} ')]", OoXmlNamespace.Manager)))
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
                    var replacementNode = mainDocumentPart.CreateElement("w", "t", OoXmlNamespace.WpML2006MainUri);
                    replacementNode.InnerText = formattedFixedDate ?? (date??DateTime.Now).ToLongDateString();
                    node.ParentNode.ReplaceChild(replacementNode, node);
                }
            }
        }
        /// <param name="wrTextOrRunNode">Either a <w:r> containing a <w:t> OR a <w:t></param>
        /// <param name="replacementText">the text to insert into <paramref name="wrTextOrRunNode"/></param>
        /// <param name="mainDocumentPart"></param>
        /// <param name="logger"></param>
        static void ReplaceInnerText(
                        this XmlNode wrTextOrRunNode, string replacementText, XmlDocument mainDocumentPart, ILogger logger)
        {
            
            if (string.IsNullOrEmpty(replacementText)) return;
            if ( (wrTextOrRunNode?.LocalName != "r" && wrTextOrRunNode?.LocalName != "t") || wrTextOrRunNode.Prefix!="w") 
            {
                logger.LogWarning("ReplaceInnerText called with a node type " + wrTextOrRunNode.Name);
            }

            var lines = replacementText.Split(NewLineSeparators, StringSplitOptions.None);
            if (lines.Length == 0)
            {
                return;
            }
            else
            {
                XmlNode ItselfOrItsInnerTextNode(XmlNode n) => n.SelectSingleNode("w:t", OoXmlNamespace.Manager) ?? n;

                var wtTextNode = ItselfOrItsInnerTextNode(wrTextOrRunNode);
                
                wtTextNode.InnerText = lines[0];
                var lastNodeWritten = wtTextNode;
                foreach (var line in lines.Skip(1))
                {
                    var nodeForLine= 
                        mainDocumentPart.CreateElement("w", "t", OoXmlNamespace.WpML2006MainUri);
                    var nodeForLinebreak=
                        mainDocumentPart.CreateElement("w", "br", OoXmlNamespace.WpML2006MainUri);
                    var textNode= mainDocumentPart.CreateTextNode(line);
                    nodeForLine.AppendChild(nodeForLinebreak);
                    nodeForLine.AppendChild(textNode);
                    wtTextNode.ParentNode.InsertAfter(nodeForLine, lastNodeWritten);
                    lastNodeWritten = nodeForLine;
                }
            }
        }
    }
}
