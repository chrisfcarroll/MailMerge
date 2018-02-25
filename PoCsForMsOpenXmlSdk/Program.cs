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
            Console.WriteLine("Press a key to continue");
            Console.Read();
        }

    }
}