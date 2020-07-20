namespace MailMerge
{
    public class OoXPath
    {
        public static string ElWithAttrMatches(string elementName, string attributeName, string patternToMatch)
        {
            return $"{elementName}[matches({attributeName},'{patternToMatch}')]";
        }

        public const string fldSimple = "w:fldSimple";
        public const string winstr = "w:instr";
        public const string wtext = "w:t";
        public const string wrun = "w:r";
        public const string MERGEFIELD = "MERGEFIELD";
    }
}
