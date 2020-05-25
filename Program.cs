using DevExpress.BarCodes;
using DevExpress.XtraRichEdit;
using DevExpress.XtraRichEdit.API.Native;
using DevExpress.XtraRichEdit.API.Native.Implementation;

using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;

namespace DxBarCodeRepro
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Hello World!");
        }

        private IEnumerable<(TemplateItem templateItem, NativeBookmark)> IterateFields(SubDocument document, Template template)
        {
            var textMarkenInVorlagen = template.Fields.ToList();
            var bookmarkCount = document.Bookmarks.OfType<NativeBookmark>().ToList().Count;

            while (bookmarkCount > 0)
            {
                var bookmarks = document.Bookmarks.OfType<NativeBookmark>().OrderBy(b => b.Range.End.ToInt()).ToList();
                foreach (var bookmark in bookmarks)
                {
                    foreach (var textMarke in template.Fields)
                    {
                        if (bookmark.Name == textMarke.Name && textMarkenInVorlagen.Contains(textMarke))
                        {
                            textMarkenInVorlagen.Remove(textMarke);
                            yield return (textMarke, bookmark);
                            goto End;
                        }
                    }
                }
            End:
                bookmarkCount--;
            }
            yield break;
        }

        private IEnumerable<SubDocument> IterateDocuments(Document doc)
        {
            yield return doc;

            foreach (var shape in doc.Shapes.Where(m => m.TextBox != null && m.TextBox.Document != null))
            {
                yield return shape.TextBox.Document;
            }

            foreach (var section in doc.Sections)
            {
                var headerDoc = section.BeginUpdateHeader();

                yield return headerDoc;

                section.EndUpdateHeader(headerDoc);

                var footerDoc = section.BeginUpdateFooter();

                yield return footerDoc;

                section.EndUpdateHeader(footerDoc);
            }
        }

        DocumentRange BuildBarcode(SubDocument document, DocumentPosition start, TemplateItem foundFeld)
        {
            if (!string.IsNullOrEmpty(foundFeld.Code))
            {
                using (var barCode = new BarCode())
                {
                    barCode.Symbology = Symbology.Code128;
                    barCode.Options.Code128.Charset = Code128CharacterSet.CharsetAuto;
                    barCode.Options.Code128.ShowCodeText = false;
                    barCode.Unit = GraphicsUnit.Point;

                    barCode.CodeText = foundFeld.Code;
                    barCode.CodeBinaryData = Encoding.UTF8.GetBytes(barCode.CodeText);

                    barCode.BackColor = Color.White;
                    barCode.ForeColor = Color.Black;

                    if (foundFeld.BarCodeHeight.HasValue)
                    {
                        barCode.ImageHeight = foundFeld.BarCodeHeight.Value;
                    }
                    if (foundFeld.BarCodeWidth.HasValue)
                    {
                        barCode.ImageWidth = foundFeld.BarCodeWidth.Value;
                    }
                    if (foundFeld.BarCodeRotation.HasValue)
                    {
                        barCode.RotationAngle = foundFeld.BarCodeRotation.Value;
                    }
                    if (foundFeld.BarCodeAutoScale.HasValue)
                    {
                        barCode.AutoSize = foundFeld.BarCodeAutoScale.Value;
                    }

                    barCode.DpiY = foundFeld.BarCodeDpiY;
                    barCode.DpiX = foundFeld.BarCodeDpiX;

                    barCode.Module = foundFeld.Module.HasValue ? foundFeld.Module.Value : 1f;

                    var img = document.Images.Insert(start, barCode.BarCodeImage);

                    if (foundFeld.BarCodeScaleY.HasValue)
                    {
                        img.ScaleY = foundFeld.BarCodeScaleY.Value;
                    }

                    if (foundFeld.BarCodeScaleX.HasValue)
                    {
                        img.ScaleX = foundFeld.BarCodeScaleX.Value;
                    }

                    return img.Range;
                }
            }
            return document.InsertText(start, string.Empty);
        }

        public void InsertItems(
          Document document,
          Template template
          )
        {
            using (var documentServer = new RichEditDocumentServer())
            {
                document.BeginUpdate();
                try
                {
                    foreach (var subDocument in IterateDocuments(document))
                    {
                        subDocument.BeginUpdate();
                        try
                        {
                            foreach (var (textMarke, bookmark) in IterateFields(subDocument, template))
                            {
                                var start = bookmark.Range.Start;
                                var len = bookmark.Range.Length;

                                var name = bookmark.Name;

                                subDocument.Bookmarks.Remove(bookmark);

                                subDocument.Replace(subDocument.CreateRange(start, len), "");

                                var range = InsertField(subDocument, start, textMarke);

                                var chars = subDocument.BeginUpdateCharacters(range);
                                if (chars.Hidden.HasValue && chars.Hidden.Value)
                                {
                                    chars.Hidden = false;
                                    chars.ForeColor = Color.Black;
                                }
                                subDocument.EndUpdateCharacters(chars);

                                var start1 = range.Start;
                                var len1 = range.Length;

                                if (len1 == 0)
                                {
                                    var bookmarkNameRange = subDocument.InsertText(start1, name);
                                    var bookmarkNameChars = subDocument.BeginUpdateCharacters(bookmarkNameRange);

                                    bookmarkNameChars.Hidden = true;
                                    bookmarkNameChars.ForeColor = Color.Blue;

                                    subDocument.EndUpdateCharacters(bookmarkNameChars);
                                    range = subDocument.CreateRange(bookmarkNameRange.Start, len1 + bookmarkNameRange.Length);
                                }

                                subDocument.Bookmarks.Create(range, name);
                            }

                            document.Fields.Update();
                        }
                        finally
                        {
                            subDocument.EndUpdate();
                        }
                    }
                    document.Fields.Update();
                }
                finally
                {
                    document.EndUpdate();
                }
            }
        }
        public DocumentRange InsertField(SubDocument document, DocumentPosition start, TemplateItem foundField)
        {
            if (foundField != null)
            {
                return BuildBarcode(document, start, foundField);
            }

            return document.InsertText(start, string.Empty);
        }

    }

    public class TemplateItem
    {
        public string Code { get; set; }

        public string Name { get; set; }
        public float? BarCodeHeight { get; set; }
        public float? BarCodeWidth { get; set; }
        public float? BarCodeRotation { get; set; }
        public bool? BarCodeAutoScale { get; set; }
        public float BarCodeDpiY { get; set; }
        public float BarCodeDpiX { get; set; }
        public double? Module { get; set; }
        public float? BarCodeScaleY { get; set; }
        public float? BarCodeScaleX { get; set; }
    }

    public class Template
    {
        public IEnumerable<TemplateItem> Fields { get; }

    }
}
