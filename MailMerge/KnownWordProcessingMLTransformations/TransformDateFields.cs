using System;
using System.Xml;
using Microsoft.Extensions.Logging;

namespace MailMerge;

public static class TransformDateFields
{
    public static readonly string[] DefaultDatesToReplace = {"DATE", "PRINTDATE", "SAVEDATE"};

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
    public static void MergeDateFields(this XmlDocument mainDocumentPart, ILogger logger, DateTime? date, string formattedFixedDate = null, string[] datesToReplace=null)
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
    public static void MergeDateFields(this XmlDocument mainDocumentPart, ILogger logger, string formattedFixedDate, string[] datesToReplace=default)
    {
        MergeDateFields(mainDocumentPart, logger, DateTime.Now, formattedFixedDate, datesToReplace);
    }
    
}