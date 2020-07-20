using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Xml;

namespace MailMerge
{
    public class OoXmlNamespace : Dictionary<string,string>
    {
        static readonly OoXmlNamespace instance = new OoXmlNamespace
        {
            {"w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
        };
        
        public static readonly ReadOnlyDictionary<string,string> Instance
            = new ReadOnlyDictionary<string, string>(instance);

        public static readonly string WpML2006MainUri = Instance["w"];

        public static readonly XmlNamespaceManager Manager =
            new Lazy<XmlNamespaceManager>(
              () => {
                      var mgr = new XmlNamespaceManager(new NameTable());
                      foreach (var pair in Instance){ mgr.AddNamespace(pair.Key, pair.Value); }
                      return mgr;
                    }).Value;
    }

}
