using System;
using System.Collections.Generic;
using IronOcr;
using System.Text.RegularExpressions;
using System.Linq;
using System.Drawing;
using System.Text;
using System.Globalization;

namespace CheckInvoices.FileOperations
{
    public static class PDF
    {
        public static object PdfRotation { get; private set; }

        public static PDFExtractedData ExtractPDFData(string pdfFilePath)
        {
            try
            {
                string client = null;
                string numberDate = null;
                string totalPayment = null;
                string codProduse = null;
                string number = null;
                string date = null;
                bool isSigned = true;

                var ocr = new IronTesseract();
                using (var input = new OcrInput())
                {
                    input.AddPdf(pdfFilePath);
                    OcrResult result = ocr.Read(input);
                    string pageText = result.Text;

                    client = ExtractClient(pageText);
                    number = ExtractNumber(pageText);
                    date = ExtractDate(pageText);
                    numberDate = ExtractNumberDate(pageText);
                    totalPayment = ExtractTotalPayment(pageText);
                    codProduse = ExtractcodProduse(pageText);
                    if (ContainsSemnaturaReprezentant(pageText))
                    {
                        isSigned = IsSignaturePresent(pdfFilePath, result);
                    }

                }

                return new PDFExtractedData
                {
                    Client = client,
                    Number = number,
                    Date = date,
                    numberDate = numberDate,
                    TotalPayment = totalPayment,
                    codProduse = codProduse,
                    isSigned = isSigned
                };
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error extracting PDF data with OCR: {ex.Message}");
            }

            return null;
        }

        public static bool ContainsSemnaturaReprezentant(string pageText)
        {
            // Remove diacritics and convert to lowercase for a case-insensitive comparison
            string normalizedText = RemoveDiacritics(pageText).ToLower();

            // Check if the normalized text contains "semnatura reprezentant"
            if (normalizedText.Contains("reprezentant"))
                return true;
            else
                return false;
        }
        public static string ExtractClient(string pageText)
        {
            var match = Regex.Match(pageText, @"Client\s*\n(.+)");
            if (match.Success)
            {
                return match.Groups[1].Value.Trim();
            }
            return "";
        }
        public static string ExtractNumber(string pageText)
        {
            string numberDate = ExtractNumberDate(pageText);
            var match = Regex.Match(numberDate, @"nr: (\d+)");
            if (match.Success)
            {
                return match.Groups[1].Value.Trim();
            }
            return "";
        }
        public static string ExtractDate(string pageText)
        {
            string numberDate = ExtractNumberDate(pageText);
            var match = Regex.Match(numberDate, @"data: (\d{2}\.\d{2}\.\d{4})");
            if (match.Success)
            {
                return match.Groups[1].Value.Trim();
            }
            return "";
        }

        public static string ExtractNumberDate(string pageText)
        {
            var match = Regex.Match(pageText, @"(?:nr:|nr\.)\s*(.*?)\n");
            if (match.Success)
            {
                return "nr" + match.Groups[1].Value.Trim().Replace("Cota TVA: 20%", "");
            }
            return "";
        }

        public static string ExtractTotalPayment(string pageText)
        {
            var match = Regex.Match(pageText, @"Total de plata\s*([\d,.]+)");
            if (match.Success)
            {
                return match.Groups[1].Value.Trim();
            }
            return "";
        }

        public static string ExtractcodProduse(string pageText)
        {
            int index = pageText.IndexOf("Denumire produse sau servicii");

            if (index >= 0)
            {
                var lines = pageText.Substring(index).Split('\n');

                List<string> values = new List<string>();

                int lineCount = 0;


                foreach (var line in lines.Skip(1)) 
                {
                    if (lineCount >= 4)
                    {
                        break; 
                    }

                    var matches = Regex.Matches(line, @"\b(XX)?\d{5,}\b");

                    foreach (Match match in matches)
                    {
                        values.Add(match.Value.Trim());
                    }

                    lineCount++;
                }
                return string.Join(";", values);
            }

            return "";
        }
        public static bool IsSignaturePresent(string pdfFilePath, OcrResult ocrResult)
        {
            try
            {
                string extractedText = RemoveDiacritics(ocrResult.Text);

                if (extractedText.Contains("reprezentant"))
                {
                    var textItem = ocrResult.Pages
                    .SelectMany(page => page.Paragraphs)
                    .FirstOrDefault(paragraph => paragraph.Text.Contains("reprezentant"));

                    if (textItem != null)
                    {
                        int x = textItem.Location.Left;
                        int y = textItem.Location.Bottom; 
                        int width = textItem.Width;
                        int height = 200;

                        Rectangle roiRect = new Rectangle(x, y, width, height);

                        using (Bitmap image = new Bitmap(pdfFilePath))
                        {
                            Bitmap roiImage = new Bitmap(image.Clone(roiRect, image.PixelFormat), roiRect.Width, roiRect.Height);


                            using (Bitmap edges = new Bitmap(roiImage.Width, roiImage.Height))
                            {
                                using (var graphics = Graphics.FromImage(edges))
                                {
                                    graphics.DrawImage(roiImage, new Rectangle(0, 0, edges.Width, edges.Height), 0, 0, roiImage.Width, roiImage.Height, GraphicsUnit.Pixel);
                                }

                                int edgeThreshold = 50; 
                                int edgeCount = 0;

                                for (int z = 0; z < edges.Width; z++)
                                {
                                    for (int w = 0; w < edges.Height; w++)
                                    {
                                        Color pixelColor = edges.GetPixel(z, w);
                                        if (pixelColor.R < edgeThreshold)
                                        {
                                            edgeCount++;
                                        }
                                    }
                                }

                                return edgeCount > 0;
                            }
                        }
                    }
                }
                return true;
            }
            catch (Exception ex)
            {
                //Console.WriteLine($"Error: {ex.Message}");
                return true;
            }
        }
        public static string RemoveDiacritics(string text)
        {
            string normalized = text.Normalize(NormalizationForm.FormD);
            StringBuilder builder = new StringBuilder();

            foreach (char c in normalized)
            {
                if (CharUnicodeInfo.GetUnicodeCategory(c) != UnicodeCategory.NonSpacingMark)
                {
                    builder.Append(c);
                }
            }

            return builder.ToString();
        }
    }
}

