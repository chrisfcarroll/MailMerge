MailMerge for docx Documents
============================

CommandLine Usage
-----------------

MailMerge inputFile1 outputFile1 [[inputFileN outputFileN]...] [ key=value[...] ]

    Settings can be read from the app-settings.json file.

    Example

    MailMerge input1.docx output1Bill.docx  FirstName=Bill  "LastName=O Reilly"



Component Usage
---------------

(outputStream, errors) = new MailMerge().Merge(outputStream, Dictionary);

        (bool, errors) = new MailMerge().Merge(inputFileName, Dictionary, outputFileName);
        

MailMerge does not use any desktop automation components, and should be suitable for serverside use. 


TODO
----
Overloads for multiline datasources: Lists, CSV files & .xmlx files.
Platform executables