using System;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;

namespace MsOpenXmlSdk
{
    class OpenXmlSdk25Examples
    {
        static void Say(object obj) => Console.WriteLine(obj);

        public static void Go(string filename)
        {
            // Create instance of OpenSettings
            OpenSettings openSettings = new OpenSettings
            {
                MarkupCompatibilityProcessSettings =
                    new MarkupCompatibilityProcessSettings(MarkupCompatibilityProcessMode.ProcessAllParts,
                        FileFormatVersions.Office2007
                    )
            };

            // Add the MarkupCompatibilityProcessSettings
            // Open the document with OpenSettings
            using (WordprocessingDocument wordDocument = WordprocessingDocument.Open(filename, true, openSettings))
            {
                Say(wordDocument.DocumentType);
                Say(wordDocument.MainDocumentPart.Document.Body.InnerText);
                SayApplicationProperties(wordDocument);
            }
        }

        public static void SayApplicationProperties(WordprocessingDocument document)
        {
            var props = document.ExtendedFilePropertiesPart.Properties;
            if (props.Company != null)Say("Company = " + props.Company.Text);
            if (props.Lines != null)Say("Lines = " + props.Lines.Text);
            if (props.Manager != null)Say("Manager = " + props.Manager.Text);

        }
    }
}