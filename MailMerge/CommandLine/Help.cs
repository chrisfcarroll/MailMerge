using System;

namespace MailMerge.CommandLine
{
    static class Help
    {
        public static void PrintHelpAndExitIfNot(bool ok)
        {
            if (ok) return;
            Console.WriteLine(HelpText);
            Environment.Exit(0);
        }

        static readonly string HelpText = 
            @"
MailMerge inputFile1 outputFile1 [[inputFileN outputFileN]...] [ key=value[...] ]

MailMerge --showxml file [fileN ...]

    Settings can be read from the app-settings.json file.

    Example

    MailMerge input1.docx output1Bill.docx  FirstName=Bill  'LastName=O Reilly'

";
    }
}