using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Xml;

namespace MailMerge.OoXml
{
    public class OoXmlNamespaces : Dictionary<string,string>
    {
        static readonly OoXmlNamespaces privateInstance = new OoXmlNamespaces
        {
            {"w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
        };
        
        public static readonly ReadOnlyDictionary<string,string> 
            Instance 
                = new ReadOnlyDictionary<string, string>(privateInstance);

        public static XmlNamespaceManager Manager =
            new Lazy<XmlNamespaceManager>(
              () => {
                      var mgr = new XmlNamespaceManager(new NameTable());
                      foreach (var pair in Instance){ mgr.AddNamespace(pair.Key, pair.Value); }
                      return mgr;
                    }).Value;
    }

}
