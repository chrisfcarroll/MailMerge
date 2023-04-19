MailMerge for docx Documents
============================

MailMerge replaces simple and complex merge fields in WordProcessingML .docx files
and helps you apply .Net's Xml & XPath tooling to Word documents.

Component Usage
---------------
#### For Streams:
```
var (outputStream, errors) = new MailMerger().Merge(inputStream, Dictionary);
```
#### For Files:
```
var (ok,errors) = new MailMerger().Merge(inputFileName, Dictionary, outputFileName);
```
Set the current DateTime for date merge fields on `MailMerger.DateTime`, e.g. `new MailMerger{DateTime=...}`

CommandLine Usage
-----------------
Merge a docx file:
```shell
dotnet MailMerge.dll inputFile1 outputFile1 [inputFileN [...outputFileN]] [ key=value [...] ]
```
- example:
```shell
dotnet MailMerge.dll MyTemplate.docx MyLetterToBill.docx  FirstName=Bill  "LastName=O Reilly"
```
Show a document's Xml:
```shell
dotnet MailMerge.dll  --showxml file [fileN ...]
```

Component Extension Methods & Helpers
---------------------------

```
stream.AsWordprocessingDocument(isEditable)
stream.AsXPathDocOfWordprocessingMainDocument(isEditable)
stream.AsXElementOfWordprocessingMainDocument(isEditable)

stream.GetXmlDocumentOfWordprocessingMainDocument()
fileInfo.GetXElementOfWordprocessingMainDocument()
fileInfo.GetXmlDocumentOfWordprocessingMainDocument()
```
A NamespaceManager, NameTable & Uri which you need when creating an XmlDocument
and/or XElements:
```
var xdoc = new XmlDocument(OoXmlNamespace.Manager.NameTable)
var xelement= mainDocumentPart.CreateElement("w", "t", OoXmlNamespace.WpML2006MainUri)
```

Settings
--------
None really, but see https://github.com/chrisfcarroll/MailMerge/blob/master/MailMerge/appsettings.json for settable limits.
For commandline --showxml usage in scripts, you'll want to set appsettings.json::Logging.LogLevel to Information or higher

Doesn't do
----------
- Anything except simple fields, complex fields, DATE, PRINTDATE & SAVEDATE fields
- Date formatting codes except a b B d D M y & h m s
- Style/Formatting codes except these Date/Time formats
- Multi-row datasources, just does 1 row at a time

Gotchas
-------
- Interprets .Net DateTime formatting codes that aren't in the WordProcessingML spec.

NuGet
-----
https://www.nuget.org/packages/MailMerge/


How do I create a Word document with merge fields without attaching a datasource?
--------------------------------------------------------

https://www.cafe-encounter.net/p2247/add-merge-fields-to-a-word-document-before-adding-a-datasource
