using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml;
using Microsoft.Extensions.Logging;

namespace MailMerge;

public static class TransformSimpleMergeFields
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
    /// <remarks>
    /// ECMA-376 Part 1 17.16.2 XML representation
    /// Fields shall be implemented in XML using either of two approaches:
    /// • As a simple field implementation, using the fldSimple element, or
    /// • As a complex field implementation, using a set of runs involving the fldChar and instrText elements.
    /// </remarks>
    public static void MergeSimpleMergeFields(this XmlDocument mainDocumentPart, Dictionary<string, string> fieldValues, ILogger log)
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
}