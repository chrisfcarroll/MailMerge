using System.Xml;

namespace MailMerge.Helpers
{
    public static class XmlNodeExtensions
    {
        /// <summary>Abbreviation for <paramref name="me"></paramref><c>.ParentNode.RemoveChild(me)</c></summary>
        /// <param name="me"></param>
        public static void RemoveMe(this XmlNode me) => me.ParentNode.RemoveChild(me);
    }
}