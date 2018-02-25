using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Office2013.Word;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace MsOpenXmlSdk
{
    class Word2010SdtExamples
    {
        public static void Go(params string[] args)
        {
            Console.WriteLine("Hello World!");

            var result = SearchAndReplace("./template-cc-tag.docx");

            if (result.UnfilledTokens.Any())
            {
                var unfilledTokens = result.UnfilledTokens;

                Console.WriteLine();
                Console.WriteLine("You have unfilled tokens in this document:");

                foreach (var unfilledToken in unfilledTokens)
                    Console.WriteLine("{0} x{1}", unfilledToken.Name, unfilledToken.Count);
            }

            Console.WriteLine();
            Console.WriteLine("Finished. Press enter to quit");
            Console.ReadLine();
        }

        public static readonly Dictionary<string, Func<string>> Mappings = new Dictionary<string, Func<string>>
        {
            {"colour", () => "red"},
            {"disposition", () => "gay"},
            {"document.generated_at", () => DateTime.UtcNow.ToString(CultureInfo.CurrentCulture)},
            //{"author.name", () => "Oliver Nixon"}
        };

        public static GenerationResult SearchAndReplace(string document)
        {
            var unfilledTokens = new UnfilfilledTokens();

            using (var templateStream = new MemoryStream())
            {
                File.Open(document, FileMode.Open).CopyTo(templateStream);
                templateStream.Position = 0;

                using (var wordDoc = WordprocessingDocument.Open(templateStream, isEditable: true))
                {
                    var mainDocument = wordDoc.MainDocumentPart;

                    //// https://stackoverflow.com/questions/31750228/replacing-text-of-content-controls-in-openxml#answer-31755783
                    var elements = mainDocument.Document.Body.Descendants<SdtElement>()
                        .Concat(mainDocument.HeaderParts.SelectMany(x => x.Header.Descendants<SdtElement>()))
                        .Concat(mainDocument.FooterParts.SelectMany(x => x.Footer.Descendants<SdtElement>()))
                        .ToList();
                    //ReplaceTokensByName(elements, unfilledTokens);
                    
                    ReplaceTokensByTag(elements, unfilledTokens);

                    wordDoc.SaveAs("./output.docx");
                }
            }

            return new GenerationResult
            {
                UnfilledTokens = unfilledTokens
            };
        }

        // https://msdn.microsoft.com/en-us/library/office/gg605189(v=office.14).aspx
        public static void ReplaceTokensByTag(IEnumerable<SdtElement> fields, UnfilfilledTokens unfilledTokens)
        {
            var tokens = fields
                .Select(x => new
                {
                    Element = x,
                    Tag = x.SdtProperties.GetFirstChild<Tag>()
                })
                .Where(x => x.Tag != null)
                .Select(x => new
                {
                    x.Element,
                    Name = x.Tag.Val.Value
                })
                .GroupBy(x => x.Name)
                .Select(g => new
                {
                    Name = g.Key,
                    Elements = g.Select(x => x.Element)
                })
                .ToList();

            foreach (var token in tokens)
            {
                Console.WriteLine("   Replacing token {0} x{1}", token.Name, token.Elements.Count());

                if (token.Name == "table1")
                {
                    // https://msdn.microsoft.com/en-us/library/cc197932(office.12).aspx
                    var sdtElement = token.Elements.First();
                    var theTable = sdtElement.Descendants<Table>().Single();
                    var theRow = theTable.Elements<TableRow>().Last();

                    foreach (var data in new[]
                    {
                        new[]{"1", "2", "3"},
                        new[]{"10", "20", "30"}
                    })
                    {
                        var rowCopy = (TableRow)theRow.CloneNode(true);
                        var cells = rowCopy.Descendants<TableCell>();
                        for (var i = 0; i < data.Length; i++)
                        {
                            cells.ElementAt(i).Append(new Paragraph(new Run(new Text(data[i]))));
                        }
                        //rowCopy.Descendants<TableCell>().ElementAt(0).Append(new Paragraph
                        //    (new Run(new Text(data.Contact.ToString()))));
                        //rowCopy.Descendants<TableCell>().ElementAt(1).Append(new Paragraph
                        //    (new Run(new Text(data.NameOfProduct.ToString()))));
                        //rowCopy.Descendants<TableCell>().ElementAt(2).Append(new Paragraph
                        //    (new Run(new Text(data.Amount.ToString()))));
                        theTable.AppendChild(rowCopy);
                    }
                    theTable.RemoveChild(theRow);
                    RemoveContentControl(sdtElement);
                    continue;
                }

                if (token.Name == "repeat1")
                {
                    // https://msdn.microsoft.com/en-us/library/cc197932(office.12).aspx
                    var sdtElement = token.Elements.First();

                    var repeat = sdtElement.Descendants<SdtRepeatedSection>().First();
                    //var theTable = sdtElement.Descendants<Table>().Single();
                    //var theRow = theTable.Elements<TableRow>().Last();

                    foreach (var data in new[]
                    {
                        new[]{"1", "2", "3"},
                        new[]{"10", "20", "30"}
                    })
                    {
                        var clone = repeat.CloneNode(true);
                        clone.InsertAfterSelf(repeat);
                            //repeat.InsertBefore()
                        //ReplaceTokensByTag(sdtElement.Descendants<SdtElement>(), unfilledTokens);
                    }
                    //theTable.RemoveChild(theRow);
                    RemoveContentControl(sdtElement);
                    continue;
                }

                if (!Mappings.TryGetValue(token.Name, out Func<string> mappedValue))
                {
                    unfilledTokens.Add(token.Name, token.Elements.Count());
                    continue;
                }

                ReplaceElementValue(token.Elements, mappedValue());
            }
        }

        public static void ReplaceTokensByName(IEnumerable<SdtElement> fields, UnfilfilledTokens unfilledTokens)
        {
            foreach (var sdtElement in fields)
            {
                var alias = sdtElement.Descendants<SdtAlias>().FirstOrDefault();
                if (alias == null) continue;

                var tokenName = alias.Val.Value;

                if (!Mappings.TryGetValue(tokenName, out Func<string> mappedValue))
                {
                    unfilledTokens.Add(tokenName);
                    continue;
                }

                ReplaceElementValue(sdtElement, mappedValue());
            }
        }

        private static void ReplaceElementValue(IEnumerable<SdtElement> sdtElements, string mappedValue)
        {
            foreach (var sdtElement in sdtElements)
            {
                ReplaceElementValue(sdtElement, mappedValue);
            }
        }

        private static void ReplaceElementValue(OpenXmlElement sdtElement, string mappedValue)
        {
            var t = sdtElement.Descendants<Text>().FirstOrDefault();

            t.Text = mappedValue;

            RemoveContentControl(sdtElement);
        }

        private static void RemoveContentControl(OpenXmlElement sdtElement)
        {
            // https://stackoverflow.com/questions/36599855/how-to-insert-text-into-a-content-control-with-the-open-xml-sdk#answer-36623224
            IEnumerable<OpenXmlElement> elements;
            if (sdtElement is SdtBlock)
                elements = (sdtElement as SdtBlock).SdtContentBlock.Elements();
            else if (sdtElement is SdtCell)
                elements = (sdtElement as SdtCell).SdtContentCell.Elements();
            else if (sdtElement is SdtRun)
                elements = (sdtElement as SdtRun).SdtContentRun.Elements();
            //else if (sdtElement is SdtRepeatedSection)
            //    elements = (sdtElement as SdtRepeatedSection).Elements();
            else
                return;

            foreach (var el in elements)
                sdtElement.InsertBeforeSelf(el.CloneNode(true));
            sdtElement.Remove();
        }

        public class GenerationResult
        {
            public UnfilfilledTokens UnfilledTokens { get; set; }
        }
    }

    public class UnfilfilledTokens : IEnumerable<UnfilfilledToken>
    {
        private readonly Dictionary<string, int> _unfilfilledTokens = new Dictionary<string, int>();

        public void Add(string token)
        {
            Add(token, 1);
        }

        public void Add(string token, int count)
        {
            if (!_unfilfilledTokens.ContainsKey(token))
            {
                _unfilfilledTokens.Add(token, 0);
            }

            _unfilfilledTokens[token] = _unfilfilledTokens[token] + count;
        }

        public IEnumerator<UnfilfilledToken> GetEnumerator()
        {
            return _unfilfilledTokens
                .Select(x => new UnfilfilledToken(x.Key, x.Value))
                .GetEnumerator();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }
    }

    public class UnfilfilledToken
    {
        public UnfilfilledToken(string name, int count)
        {
            Name = name;
            Count = count;
        }

        public string Name { get; }
        public int Count { get; }
    }
}