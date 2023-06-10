using System;
using System.Linq;
using System.Xml;
using Microsoft.Extensions.Logging;

namespace MailMerge
{
    public static class XmlNodeTextExtensions
    {
        internal static readonly string[] NewLineSeparators =  {"\n", "\n\r"};

        /// <param name="wrTextOrRunNode">Either a <w:r> containing a <w:t> OR a <w:t></param>
        /// <param name="replacementText">the text to insert into <paramref name="wrTextOrRunNode"/></param>
        /// <param name="mainDocumentPart"></param>
        /// <param name="logger"></param>
        public static void ReplaceInnerText(
                        this XmlNode wrTextOrRunNode, string replacementText, XmlDocument mainDocumentPart, ILogger logger)
        {
            
            if (replacementText == null) return;
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
