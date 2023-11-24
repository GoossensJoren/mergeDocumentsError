using System.Collections.Generic;
using System.IO;
using System.Linq;
using GemBox.Document;

namespace ConversionError
{
    class Program
    {
        static void Main(string[] args)
        {
            ComponentInfo.SetLicense("FREE-LIMITED-KEY");

            var document1 = DocumentModel.Load(Path.Combine(Directory.GetCurrentDirectory(), "file1.docx"));
            var document2 = DocumentModel.Load(Path.Combine(Directory.GetCurrentDirectory(), "file2.docx"));
            var document3 = DocumentModel.Load(Path.Combine(Directory.GetCurrentDirectory(), "file3.docx"));
            var document4 = DocumentModel.Load(Path.Combine(Directory.GetCurrentDirectory(), "file4.docx"));
            var document5 = DocumentModel.Load(Path.Combine(Directory.GetCurrentDirectory(), "file5.docx"));
            var document6 = DocumentModel.Load(Path.Combine(Directory.GetCurrentDirectory(), "file6.docx"));

            var mergedDocument = MergeDocuments(new List<DocumentModel> { document1, document2, document3, document4, document5, document6 }, false);

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
                    importedSection.PageSetup.SectionStart = SectionStart.NewPage;
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
    }
