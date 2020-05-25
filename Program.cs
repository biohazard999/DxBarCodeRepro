using DevExpress.BarCodes;
using DevExpress.XtraRichEdit.API.Native;

using System;
using System.Drawing;
using System.Text;

namespace DxBarCodeRepro
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Hello World!");
        }

        DocumentRange BuildBarcode(SubDocument document, string code, DocumentPosition start, BarcodeOptions foundFeld)
        {
            if (!string.IsNullOrEmpty(code))
            {
                using (var barCode = new BarCode())
                {
                    barCode.Symbology = Symbology.Code128;
                    barCode.Options.Code128.Charset = Code128CharacterSet.CharsetAuto;
                    barCode.Options.Code128.ShowCodeText = false;
                    barCode.Unit = GraphicsUnit.Point;

                    barCode.CodeText = code;
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
    }

    public class BarcodeOptions
    {
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
}
