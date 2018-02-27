using System.ComponentModel;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

[assembly: InternalsVisibleTo("MailMerge.OoXml.Tests")]
[assembly: AssemblyConfiguration("")]
[assembly: AssemblyCompany("Chris F Carroll")]
[assembly: AssemblyProduct("MailMerge")]
[assembly: ComVisible(false)]
[assembly: Guid("d1c4ab83-c553-4e3b-8e75-c9e76498206b")]

[assembly: AssemblyDescription(@"MailMerge for docx Documents
============================

CommandLine Usage
-----------------

MailMerge inputFile1 outputFile1 [[inputFileN outputFileN]...] [ key=value[...] ]

Settings can be read from the app-settings.json file.

Example

MailMerge input1.docx output1Bill.docx  FirstName=Bill  ""LastName=O Reilly""

Component Usage
---------------

(outputStream, errors) = new MailMerge().Merge(outputStream, Dictionary);

(bool, errors) = new MailMerge().Merge(inputFileName, Dictionary, outputFileName);
        
TODO
----
Overloads for multiline datasources: Lists, CSV files & .xmlx files.
Platform executables
")]