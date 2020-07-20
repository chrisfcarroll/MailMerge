using System;
using System.Collections.Generic;
using System.IO;
using System.Xml;
using Extensions.Logging.ListOfString;
using MailMerge.CommandLine;
using NUnit.Framework;
using TestBase;

namespace MailMerge.Tests.FunctionalSpecs
{
    [TestFixture]
    public class GivenADateMergeInstruction
    {
        MailMerger sut;

        [TestCase(dateFragmentWithddMMMMyyyyFormat)]
        public void ReplacesWithTheDate(string fragment)
        {
            const string myFormattedDate = "thisismydate";
            var xdoc = new XmlDocument(OoXmlNamespace.Manager.NameTable);
            xdoc.LoadXml(fragment);
            xdoc.MergeDate(new StringListLogger(), myFormattedDate);
            
            Console.WriteLine(xdoc.OuterXml);
            xdoc.OuterXml.Replace("\n", "").Replace("\r", "").ShouldMatch($@"\<w:r\>\s*\<w:t\>{myFormattedDate}\</w:t\>\s*\</w:r\>");
        }

        [TestCase("dd MMMM yyyy",dateFragmentWithddMMMMyyyyFormat)]
        [TestCase("dd-MM-yy",dateFragmentWithddMMyyFormat)]
        public void ReplacesWithTheDateAndRespectsFormat(string format, string fragment)
        {
            var date = DateTime.Now;
            var xdoc = new XmlDocument(OoXmlNamespace.Manager.NameTable);
            xdoc.LoadXml(fragment);

            xdoc.MergeDate(new StringListLogger(),date);
            
            Console.WriteLine(xdoc.OuterXml);
            xdoc.OuterXml.Replace("\n", "").Replace("\r", "").ShouldMatch($@"\<w:r\>\s*\<w:t\>{date.ToString(format)}\</w:t\>\s*\</w:r\>");
        }


        [Test]
        public void ReplacesWithGivenFieldValuesDate()
        {
            using(var instream= new FileStream(Path.Combine("TestDocuments","ATemplate.docx"),FileMode.Open))
            using (var outStream = new MemoryStream( new byte[instream.Length * 2] ))
            {
                const string myFormattedDate = "thisismydate";
                instream.CopyTo(outStream);

                sut.ApplyAllKnownMergeTransformationsToMainDocumentPart(new Dictionary<string, string>{{MailMerger.DATEKey,myFormattedDate}}, outStream);
                
                outStream.Position = 0;
                var output = outStream.AsXElementOfWordprocessingMainDocument().ToString();
                output.Replace("\n", "").Replace("\r", "").ShouldMatch($@"\<w:r\>\s*\<w:t\>{myFormattedDate}\</w:t\>\s*\</w:r\>");
            }
        }

        [SetUp]
        public void Setup()
        {
            sut = new MailMerger(new StringListLogger() , Startup.Settings);
        }

        const string dateFragmentWithddMMMMyyyyFormat = @"<?xml version='1.0' encoding='UTF-8' ?>
<w:document xmlns:w='http://schemas.openxmlformats.org/wordprocessingml/2006/main'>
	<w:body>
		<w:p>
			<w:r>
				<w:instrText xml:space='preserve'>DATE  \@ &quot;dd MMMM yyyy&quot;  \* MERGEFORMAT </w:instrText>
			</w:r>
		</w:p>
	</w:body>
</w:document>";

        const string dateFragmentWithddMMyyFormat = @"<?xml version='1.0' encoding='UTF-8' ?>
<w:document xmlns:w='http://schemas.openxmlformats.org/wordprocessingml/2006/main'>
	<w:body>
		<w:p>
			<w:r>
				<w:instrText xml:space='preserve'>DATE  \@ &quot;dd-MM-yy&quot;  \* MERGEFORMAT </w:instrText>
			</w:r>
		</w:p>
	</w:body>
</w:document>";

    }
}
