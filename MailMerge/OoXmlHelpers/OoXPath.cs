namespace MailMerge
{
    /// <summary>
    /// Some example Constants and a Helper Method for using XPath on OOXml
    /// </summary>
    public static class OoXPath
    {
        /// <summary>Name of an element matching ECMA-376 Part 1 17.16.5.35
        /// <![CDATA[<w:fldSimple w:instr=" MERGEFIELD Name "><w:t>«Name»</w:t></w:fldSimple>]]>  
        /// </summary>
        public const string fldSimple = "w:fldSimple";
        
        /// <summary>
        /// Name of an instuction attribute used in e.g. ECMA-376 Part 1 17.16.5.35
        /// <![CDATA[<w:fldSimple w:instr=" MERGEFIELD Name "><w:t>«Name»</w:t></w:fldSimple>]]>  
        /// </summary>
        public const string winstr = "w:instr";
        
        /// <summary>Name of a text element : ECMA-376 Part 1 17.3.3.31 t (Text)</summary>
        public const string wtext = "w:t";
        
        /// <summary>Name of a Run element : ECMA-376 Part 1 17.3.2</summary>
        public const string wrun = "w:r";
        
        /// <summary>String value of <see cref="winstr"/> which identifies a MergeField instruction</summary>
        public const string MERGEFIELD = "MERGEFIELD";
        
        /// <summary>Generate an XPath snippet to match elements of name
        /// <paramref name="elementName"/> with attributes of name <paramref name="attributeName"/>
        /// whose values match <paramref name="patternToMatch"/>
        /// </summary>
        /// <param name="elementName"></param>
        /// <param name="attributeName"></param>
        /// <param name="patternToMatch">the pattern which each <c>element.attribute</c> should match</param>
        /// <returns><c>$"{elementName}[matches({attributeName},'{patternToMatch}')]"</c></returns>
        public static string ElWithAttrMatches(string elementName, string attributeName, string patternToMatch)
        {
            return $"{elementName}[matches({attributeName},'{patternToMatch}')]";
        }
    }
}
