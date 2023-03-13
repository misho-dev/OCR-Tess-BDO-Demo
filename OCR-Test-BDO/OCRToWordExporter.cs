using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Diagnostics;
using Tesseract;
using Xceed.Words.NET;

namespace OCR_Test_BDO
{
    internal class OCRToWordExporter
    {
        public static void ExportToWord(string imagePath, string wordPath)
        {
            Environment.SetEnvironmentVariable("TESSDATA_PREFIX", "C:/Users/m.saamashvili/source/repos/OCR-Test-BDO/tessdata");

            // Check if the input file exists
            if (!File.Exists(imagePath))
            {
                throw new ArgumentException($"Input file {imagePath} does not exist.");
            }

            try
            {
                // Pre-process the image
                var preprocessedImagePath = PreprocessImage(imagePath);


                // Create a Tesseract OCR engine
                using (var engine = new TesseractEngine("C:/Users/m.saamashvili/source/repos/OCR-Test-BDO/tessdata", "eng", EngineMode.Default))
                {
                    // Load the pre-processed image
                    using (var img = Pix.LoadFromFile(preprocessedImagePath))
                    {
                        // Set the page segmentation mode to Automatic
                        engine.SetVariable("tessedit_pageseg_mode", "auto");

                        // Recognize text from the image
                        using (var page = engine.Process(img))
                        {
                            // Get the recognized text
                            var text = page.GetText();

                            // Create a Word document
                            var doc = DocX.Create(wordPath);

                            // Add the recognized text to the document
                            doc.InsertParagraph(text);

                            // Save the document
                            doc.Save();
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                throw new Exception($"Error processing file {imagePath}: {ex.Message}");
            }
        }

        private static string PreprocessImage(string imagePath)
        {
            try
            {
                // Load the image
                using (var image = Pix.LoadFromFile(imagePath))
                {
                    // Convert the image to grayscale
                    using (var grayscaleImage = image.ConvertRGBToGray())
                    {
                        // Create a new image file name with "-processed" appended to the original file name
                        var processedImagePath = Path.Combine(Path.GetDirectoryName(imagePath), Path.GetFileNameWithoutExtension(imagePath) + "-processed" + Path.GetExtension(imagePath));

                        // Save the pre-processed image to the new file
                        grayscaleImage.Save(processedImagePath, ImageFormat.Tiff);

                        return processedImagePath;
                    }
                }
            }
            catch (Exception ex)
            {
                throw new Exception($"Error preprocessing image {imagePath}: {ex.Message}");
            }
        }
    }
}
