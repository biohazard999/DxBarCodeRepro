using DevExpress.XtraRichEdit;
using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;

namespace DxBarCodeRepro
{
    public enum DetectedDocumentFormat
    {
        OpenXml,
        Word,
        Unknown
    }

    public static class DocumentServerExtentions
    {
        private static Dictionary<DetectedDocumentFormat, DocumentFormat> Mapping = new Dictionary<DetectedDocumentFormat, DocumentFormat>
        {
            [DetectedDocumentFormat.OpenXml] = DocumentFormat.OpenXml,
            [DetectedDocumentFormat.Word] = DocumentFormat.Doc,
            [DetectedDocumentFormat.Unknown] = DocumentFormat.Undefined,
        };

        public static DetectedDocumentFormat DetectWordFormat(string path)
        {
            var ext = System.IO.Path.GetExtension(path);


            try
            {
                using (ZipArchive archive = ZipFile.Open(path, ZipArchiveMode.Update))
                {
                    return DetectedDocumentFormat.OpenXml;
                }
            }
            catch(InvalidDataException)
            {
                if (new string[] { ".doc", ".dot" }.Contains(ext))
                {
                    return DetectedDocumentFormat.Word;
                }
                else
                {
                    return DetectedDocumentFormat.Unknown;
                }
            }
        }

        public static bool LoadFileInDetectionMode(this RichEditDocumentServer server, string path)
        {
            var detectedFormat = DetectWordFormat(path);
            return LoadFileInDetectionMode(server, path, detectedFormat);
        }

        public static bool LoadFileInDetectionMode(this RichEditDocumentServer server, string path, DetectedDocumentFormat format)
        {
            var ext = System.IO.Path.GetExtension(path);
            var docFormat = Mapping[format];


            if (new string[] { ".dot", ".dotx" }.Contains(ext))
            {
                var documentLoaded = server.LoadDocumentTemplate(path, docFormat); // Dot, Dotx
                if (documentLoaded)
                {
                    server.Document.DocumentProperties.Created = System.DateTime.Now;
                }
                return documentLoaded;
            }
            else if (new string[] { ".doc", ".docx" }.Contains(ext))
            {
                return server.LoadDocument(path, docFormat); // Doc, Docx
            }
            else
            {
                return false; // Unidentified
            }

        }

    }
}
