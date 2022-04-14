using System;

namespace MsOpenXmlSdk
{
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