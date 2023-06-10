using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml;
using Microsoft.Extensions.Logging;

namespace MailMerge;

public static class TransformComplexMergeFields
{
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
    public static void MergeComplexMergeFields(this XmlDocument mainDocumentPart, Dictionary<string, string> fieldValues, ILogger log)
    {
        const int boilerplateCount = 3; 

        var beginRuns= mainDocumentPart.SelectNodes("//w:r[w:fldChar/@w:fldCharType='begin']", OoXmlNamespace.Manager);
        log.LogDebug("Found " + beginRuns.Count + " <w:fldChar w:fldCharType='begin'> nodes.");
        foreach(XmlNode beginRun in beginRuns)
        {
            var boilerPlateNodes= new List<XmlNode> { beginRun };
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

                    var fieldName = Enumerable.Skip<string>(nodesText
                        .Split(" ", StringSplitOptions.RemoveEmptyEntries), 1).FirstOrDefault();
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
    static string CollectSequentialInstrRuns(XmlNode instrNode, ref XmlNode sibling, List<XmlNode> instrRuns)
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
}