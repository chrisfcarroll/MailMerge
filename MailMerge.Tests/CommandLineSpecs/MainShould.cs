using System.IO;
using MailMerge.CommandLine;
using NUnit.Framework;

namespace MailMerge.Tests.CommandLineSpecs
{
    public class MainShould
    {
        [Test]
        public void NotThrowIfThereIsNoAppSettingsJson()
        {
            if (!File.Exists("appsettings.json"))
            {
                var thereWasNoAppsettingsIn = "There was no appsettings in " + Directory.GetCurrentDirectory();
                NUnit.Framework.Assert.Inconclusive(thereWasNoAppsettingsIn);
            }
            System.IO.File.Delete("appsettings.json.bak");
            System.IO.File.Move("appSettings.json", "appsettings.json.bak");
            
            Program.Main("in.docx", "out.docx", "a=b");
            
            System.IO.File.Move("appsettings.json.bak","appsettings.json");
        }
    }
}
