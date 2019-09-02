using iTextSharp.text;
using System;
using System.IO;

namespace TiffConverter
{
    //Using the http://pdfsharp.com/PDFsharp/ reference
    class Program
    {
        /// <summary>Defines the entry point of the application.</summary>
        /// <param name="args">The arguments.</param>
        /// <exception cref="System.Exception">Source folder path is empty!!!
        /// or
        /// Destination folder path is empty!!!</exception>
        static void Main(string[] args)
        {
            try
            {
                Console.WriteLine("Please enter the source folder path:");
                string tifSource = Console.ReadLine();
                if (tifSource == string.Empty)
                {
                    throw new System.Exception("Source folder path is empty!!!");
                }
                Console.WriteLine("Please enter the destination folder path:");
                string tifDestination = Console.ReadLine();
                if (tifDestination == string.Empty)
                {
                    throw new System.Exception("Destination folder path is empty!!!");
                }

                ConvertMultiple(tifSource, tifDestination);
            }
            catch (Exception ex)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine(ex.Message);
                Console.ForegroundColor = ConsoleColor.Gray;
            }
        }

        /// <summary>Converts the Tif files into a multi page PDF multiple.</summary>
        /// <param name="sourceDir">The source dir.</param>
        /// <param name="destinationDir">The destination dir.</param>
        /// <exception cref="System.Exception">Source folder '"+ sourceDir + "' contains no files!!!</exception>
        private static void ConvertMultiple(string sourceDir, string destinationDir)
        {
            try
            {
                DirectoryInfo di = new DirectoryInfo(sourceDir);

                int totalFiles = di.GetFiles().Length;
                if (totalFiles == 0)
                {
                    throw new System.Exception("Source folder '"+ sourceDir + "' contains no files!!!");
                }
                Console.WriteLine("Total Files in Source Folder = " + totalFiles);

                foreach (var file in di.GetFiles())
                {
                    totalFiles = totalFiles -= 1;

                    if (file.Extension.ToString() == ".tif" || file.Extension.ToString() == ".tiff")
                    {
                        iTextSharp.text.Document document = new iTextSharp.text.Document(iTextSharp.text.PageSize.A4, 0, 0, 0, 0);
                        string fileName = Path.GetFileNameWithoutExtension(file.ToString());
                        string filePath = string.Format("{0}\\{1}.{2}", destinationDir, fileName, "pdf");
                        string sourceFilePath = string.Format("{0}\\{1}.{2}", sourceDir, fileName, "tif");
                        FileStream fs = new FileStream(filePath, FileMode.Create, FileAccess.ReadWrite);
                        iTextSharp.text.pdf.PdfWriter writer = iTextSharp.text.pdf.PdfWriter.GetInstance(document, fs);

                        // Counts the files remaining to be converting 
                        Console.WriteLine("Converting: " + file.Name + " Total Files Left: " +totalFiles);

                        System.Drawing.Bitmap bmp = new System.Drawing.Bitmap(sourceFilePath);

                        document.Open();
                        iTextSharp.text.pdf.PdfContentByte cb = writer.DirectContent;

                        int total = bmp.GetFrameCount(System.Drawing.Imaging.FrameDimension.Page);
                        for (int k = 0; k < total; ++k)
                        {
                            bmp.SelectActiveFrame(System.Drawing.Imaging.FrameDimension.Page, k);
                            iTextSharp.text.Image img = iTextSharp.text.Image.GetInstance(bmp, System.Drawing.Imaging.ImageFormat.Bmp);

                            // Scale the image to fit in the page
                            img.ScalePercent(72f / img.DpiX * 100);
                            img.SetAbsolutePosition(0, 0);
                            cb.AddImage(img);
                            document.NewPage();
                        }

                        bmp.Dispose();
                        writer.Flush();
                        document.Close();
                        document.Dispose();
                    }
                    else
                    {
                        Console.WriteLine("Not Converting: " + file.Name + " Total Files Left: " + totalFiles);
                    }
                }
            }
            catch (Exception ex)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine(ex.Message);
                Console.ForegroundColor = ConsoleColor.Gray;
            }
        }
    }
}
