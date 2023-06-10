using System;

namespace MsOpenXmlSdk
{
    /// <summary>
    /// This project is largely microsoft code intended to demo the OpenXml Sdk 2.5
    /// Debug-Run it in your IDE to step through the code in <see cref="OpenXmlSdk25Examples"/>.
    /// </summary>
    class Program
    {
        static void Main(string[] args)
        {
            var filename = args.Length>0 ? args[0] : @"ATemplate.docx";

            OpenXmlSdk25Examples.Go(filename);
            //Word2010SdtExamples.Go(filename);
            Console.WriteLine("\n-------------------------------\n" +
                              "This project is largely Microsoft code intended to demo the OpenXml Sdk 25.\n\n" +
                              "Debug this in your IDE to step through basic \n" +
                              "OpenXmlSdk usage in OpenXmlSdk25Examples.cs");
            Console.WriteLine("Press the Enter key to continue");
            Console.Read();
        }

    }
}