using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using iTextSharp.text;
using iTextSharp.text.pdf;
using iTextSharp.text.pdf.parser;

namespace PDFExploder
{
    class Program
    {
        static void Main(string[] args)
        {
            StringBuilder textstuff = new StringBuilder();

            using (var pdfReader = new PdfReader("c:\\temp\\reviews2014B.pdf"))
            {
                int startpage = 1, endpage = 0;
                string datestring = string.Empty;
                PdfDocument outputDocument = new PdfDocument();
                string outputName = string.Empty;
                string name = string.Empty;

                // Loop through each page of the document
                for (var page = 1; page <= pdfReader.NumberOfPages; page++)
                {
                    ITextExtractionStrategy strategy = new SimpleTextExtractionStrategy();

                    var currentText = PdfTextExtractor.GetTextFromPage(
                    pdfReader,
                    page,
                    strategy);

                    currentText =
                        Encoding.UTF8.GetString(Encoding.Convert(
                            Encoding.Default,
                            Encoding.UTF8,
                            Encoding.Default.GetBytes(currentText)));



                    /*  if (startpage == endpage -1000000)
                      {
                         // datestring = GetDatePos(currentText);

                          //Get the Employee info
                          int startidx = currentText.LastIndexOf("Review") + 9;
                          int endidx = currentText.IndexOf("Page ", startidx) - 1;
                          int nameLength = endidx - startidx;

                          name = currentText.Substring(startidx, nameLength);
                          outputName = string.Format("c:\\temp\\{0}__{1}.pdf", name, datestring).Replace(" ", "");

                      } */




                    //The first page - need to pull the date off this
                    //This only runs the first time
                    if (startpage == 1 && endpage == 0)
                    {
                        datestring = GetDatePos(currentText);
                        outputName = string.Format("c:\\temp\\{0}__{1}.pdf", name, datestring).Replace(" ", "");
                        //Get the Employee info
                        int startidx = currentText.LastIndexOf("Review") + 9;
                        int endidx = currentText.IndexOf("Page ", startidx) - 1;
                        int nameLength = endidx - startidx;

                        name = currentText.Substring(startidx, nameLength);
                        outputName = string.Format("c:\\temp\\{0}__{1}.pdf", name, datestring).Replace(" ", "");
                    }

                    if (currentText.Contains("Instructions for the Reviewer") && endpage > startpage)
                    {
                        Console.WriteLine(string.Format("Found new review at page {0} - Extracting...", page.ToString()));


                        ////Get the Employee info
                        //int startidx = currentText.LastIndexOf("Review") + 9;
                        //int endidx = currentText.IndexOf("Page ",startidx) - 1;
                        //int nameLength = endidx - startidx;

                        //string name = currentText.Substring(startidx, nameLength);
                        //string ouputName = string.Format("c:\\temp\\{0}__{1}.pdf",name,datestring).Replace(" ","");

                        ExtractPages("c:\\temp\\reviews2014B.pdf", outputName, startpage, endpage);
                        Console.WriteLine("File extracted!");

                        //  endpage--;
                        startpage = endpage + 1;

                        //Get next date string
                        //int datestartidx = currentText.IndexOf("/") - 2;
                        datestring = GetDatePos(currentText); // Convert.ToDateTime(currentText.Substring(datestartidx, 8).TrimStart()).ToShortDateString().Replace("/", "_");

                        int startidx = currentText.LastIndexOf("Review") + 9;
                        int endidx = currentText.IndexOf("Page ", startidx) - 1;
                        int nameLength = endidx - startidx;

                        name = currentText.Substring(startidx, nameLength);
                        outputName = string.Format("c:\\temp\\{0}__{1}.pdf", name, datestring).Replace(" ", "");

                    }

                    endpage++;
                    // textstuff.Append(currentText);
                }
            }

            Console.ReadLine();
        }

        static string GetDatePos(string currentText)
        {
            string thedate = string.Empty;
            /***** START OF REVIEW SAMPLE *****/
            //Associate Performance Review
            //Employee Name Title Evaluation Date
            //Abhalter, Karen R Payroll Techician 2/27/15
            //Department Manager Date of Hire
            int theDigit;
            int startidx;

            int firstSlashidx = currentText.IndexOf("/");

            if (int.TryParse(currentText[firstSlashidx - 1].ToString(), out theDigit))
            {
                startidx = firstSlashidx - 2;
            }
            else
            {
                startidx = currentText.IndexOf("/", firstSlashidx + 1) - 2;
            }

            thedate = Convert.ToDateTime(currentText.Substring(startidx, 8).TrimStart()).ToShortDateString().Replace("/", "_");

            return thedate;


        }
        static void ExtractPages(string sourcePdfPath, string outputPdfPath,
        int startPage, int endPage)
        {
            PdfReader reader = null;
            Document sourceDocument = null;
            PdfCopy pdfCopyProvider = null;
            PdfImportedPage importedPage = null;

            try
            {
                // Intialize a new PdfReader instance with the contents of the source Pdf file:
                reader = new PdfReader(sourcePdfPath);

                // For simplicity, I am assuming all the pages share the same size
                // and rotation as the first page:
                sourceDocument = new Document(reader.GetPageSizeWithRotation(startPage));

                // Initialize an instance of the PdfCopyClass with the source 
                // document and an output file stream:
                pdfCopyProvider = new PdfCopy(sourceDocument,
                    new System.IO.FileStream(outputPdfPath, System.IO.FileMode.Create));

                sourceDocument.Open();

                // Walk the specified range and add the page copies to the output file:
                for (int i = startPage; i <= endPage; i++)
                {
                    importedPage = pdfCopyProvider.GetImportedPage(reader, i);
                    pdfCopyProvider.AddPage(importedPage);
                }
                sourceDocument.Close();
                reader.Close();
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
    }
}
