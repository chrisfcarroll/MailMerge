MailMerge for docx Documents
============================

MailMerge replaces simple and complex merge fields in WordProcessingML .docx files.

Component Usage
---------------
```
(outputStream, errors) = new MailMerger().Merge(inputStream, Dictionary);
```
or
```
new MailMerger{DateTime=...}.Merge(inputFileName, Dictionary, outputFileName);
```

CommandLine Usage
-----------------
```
dotnet MailMerge.dll inputFile1 outputFile1 [inputFileN [...outputFileN]] [ key=value[...] ]
```

Example
```
dotnet MailMerge.dll input1.docx output1Bill.docx  FirstName=Bill  "LastName=O Reilly"
```


Doesn't do
----------
- Anything except Merge fields and Dates
- Overloads for multiline datasources
