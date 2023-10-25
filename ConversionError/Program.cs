using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using GemBox.Document;
using GemBox.Document.Tables;

namespace ConversionError
{
    class Program
    {
        static void Main(string[] args)
        {
            ComponentInfo.SetLicense("FREE-LIMITED-KEY");

            var document1 = DocumentModel.Load(Path.Combine(Directory.GetCurrentDirectory(), "file1.docx"));

            var mergedDocument = MergeDocuments(new List<DocumentModel> { document1 }, false);

            using var stream = new MemoryStream();
            mergedDocument.Save("merged.docx", SaveOptions.DocxDefault);
        }

        public static DocumentModel MergeDocuments(List<DocumentModel> documents, bool usePageNumbering)
        {
            var mergedDocument = new DocumentModel();
            var first = true;
            DocumentModel lastDocument = documents.Last();

            foreach (DocumentModel document in documents)
            {
                // Make sure no styling changes occur
                RenameStylesToBeUnique(document, documents.IndexOf(document));

                var mapping = new ImportMapping(document, mergedDocument, false);

                if (document.Styles.Contains("Hyperlink") && mergedDocument.Styles.Contains("Hyperlink"))
                {
                    mapping.SetDestinationStyle(document.Styles["Hyperlink"], mergedDocument.Styles["Hyperlink"]);
                }

                if (document != lastDocument)
                {
                    AddEmptyPage(document);
                }

                SectionCollection documentSections = document.Sections;

                Section firstSection = document.Sections.First();
                if (firstSection != null && usePageNumbering)
                {
                    firstSection.PageSetup.PageStartingNumber = 1;
                }

                foreach (Section section in documentSections)
                {
                    Section importedSection = mergedDocument.Import(section, true, mapping);
                    mergedDocument.Sections.Add(importedSection);
                }

                if (!first)
                {
                    continue;
                }

                mergedDocument.DefaultCharacterFormat = document.DefaultCharacterFormat.Clone();
                mergedDocument.DefaultParagraphFormat = document.DefaultParagraphFormat.Clone();
                first = false;
            }

            return mergedDocument;
        }

        private static void AddEmptyPage(DocumentModel document)
        {
            int pageAmount = document.GetPaginator().Pages.Count;

            if ((pageAmount % 2) != 0)
            {
                document.Sections.Add(new Section(document));
            }
        }

        private static void RenameStylesToBeUnique(DocumentModel document, long index)
        {
            // Every document has a default style (mostly) called "Normal" but this style can differentiate from document to document
            // So rename all the styles to be unique (by adding the index of the document)
            List<Style> stylesOfCurrentDocument = document.Styles.ToList();

            foreach (Style style in stylesOfCurrentDocument)
            {
                List<Block> elementsInDocumentWithStyle = document.Sections.SelectMany(s => s.Blocks.ToList())
                    .SelectMany(b => b is Table table ? new List<Block> { table }.Concat(table.Rows.SelectMany(r => r.Cells.SelectMany(c => c.Blocks))) : new List<Block> { b })
                    .Where(b => (b is Paragraph paragraph && paragraph.ParagraphFormat.Style == style) || (b is Table table && table.TableFormat.Style == style))
                    .ToList();

                style.Name += $"_{index}";

                document.Save("test.xml");

                foreach (Block block in elementsInDocumentWithStyle)
                {
                    switch (block)
                    {
                        case Paragraph paragraph when paragraph.ParagraphFormat.Style != null:
                            paragraph.ParagraphFormat.Style.Name = style.Name;
                            //ClearStylingIfConflicting(paragraph);
                            break;
                        case Table table:
                            table.TableFormat.Style.Name = style.Name;
                            break;

                    }
                }
            }
        }

        private static void ClearStylingIfConflicting(Paragraph paragraph)
        {
            ParagraphFormat originalParagraphFormat = paragraph.ParagraphFormat.Clone();
            ParagraphFormat styleParagraphFormat = paragraph.ParagraphFormat.Style?.ParagraphFormat.Clone();

            if (styleParagraphFormat == null)
            {
                return;
            }

            if (Math.Abs(originalParagraphFormat.SpaceAfter - styleParagraphFormat.SpaceAfter) < 0.0000001 &&
                Math.Abs(originalParagraphFormat.SpaceBefore - styleParagraphFormat.SpaceBefore) < 0.0000001 &&
                Math.Abs(originalParagraphFormat.LeftIndentation - styleParagraphFormat.LeftIndentation) < 0.0000001 &&
                Math.Abs(originalParagraphFormat.RightIndentation - styleParagraphFormat.RightIndentation) < 0.0000001)
            {
                return;
            }

            paragraph.ParagraphFormat.Style = null;
        }
    }
}
