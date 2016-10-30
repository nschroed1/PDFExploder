using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using iTextSharp.text;
using iTextSharp.text.pdf;
using iTextSharp.text.pdf.parser;
using System.IO;

namespace PDFExploder
{
    class Program
    {

        static void OutputSlip(string thePayslip, string empName, string header)
        {
            string[] nameParts;
            nameParts = empName.Split(',');
            string filePath = string.Format("c:\\temp\\payslips\\{0}_{1}.txt",nameParts[1],nameParts[0]);
            string pdfPath = string.Format("c:\\temp\\payslips\\{0}_{1}.pdf", nameParts[1], nameParts[0]);
            TextWriter importFile = new StreamWriter(filePath);
            importFile.Write(header);
            importFile.WriteLine();
            importFile.Write(thePayslip);
            importFile.Close();

            //Read the Data from Input File
            StreamReader rdr = new StreamReader(filePath);
            //Create a New instance on Document Class
            Document doc = new Document();
            //Create a New instance of PDFWriter Class for Output File
            PdfWriter.GetInstance(doc, new FileStream(pdfPath, FileMode.Create));
            //Open the Document
            doc.Open();
            Font theFont = new Font(Font.FontFamily.COURIER, 8);
            //Add the content of Text File to PDF File
            doc.Add(new Paragraph(rdr.ReadToEnd(),theFont));
            //Close the Document
            doc.Close();

        }

        static void ExtractPayslips()
        {
            //Read Files in the directory
            DirectoryInfo di = new DirectoryInfo("C:\\temp\\W2s and Payslips\\Payslips\\2015");
            FileInfo [] theFiles = di.GetFiles();
            TextReader tr = null;
            string currentLine = string.Empty;
            StringBuilder payslipstring = new StringBuilder();
            string empName = string.Empty;
            bool gotName = false;
            StringBuilder headerLine = new StringBuilder();
            bool gotHeader = false;

            foreach (FileInfo fi in theFiles )
            {
                using (tr = new StreamReader(fi.FullName))
                {
                    while (tr.Peek() != -1)
                    { 
                        currentLine = tr.ReadLine();

                        if ( !gotHeader && currentLine.Contains("EARNINGS"))
                        {
                            headerLine.AppendLine(currentLine);
                        }

                        if (!gotHeader && currentLine.Contains("CONTRIBUTIONS") || gotHeader)
                        {
                            if (!gotHeader)
                            {
                                headerLine.AppendLine(currentLine);
                                gotHeader = true;
                            }



                            //this should be a blank line
                            if (currentLine.Contains("CONTRIBUTIONS"))
                               currentLine = tr.ReadLine();

                            gotName = false;

                            //Now we are getting data
                            while (!currentLine.Contains("TOTAL"))
                            {
                                currentLine = tr.ReadLine();

                                if (!gotName)
                                {
                                    int nameStart = currentLine.IndexOf(" ");
                                    int nameEnd = currentLine.IndexOf("      ");
                                    empName = currentLine.Substring(nameStart, nameEnd - nameStart);
                                    empName.Trim();
                                    
                                    gotName = true;
                                }

                                //remove continued verbiage
                                if (currentLine.Contains("CONTINUED"))
                                {
                                    int x = 1;
                                    //currentLine = currentLine

                                    int end = currentLine.IndexOf("CONTINUED **") + 12;
                                    string removeText = currentLine.Substring(0, end);
                                    string newText = " ";
                                    newText.PadRight(removeText.Length, ' ');

                                    currentLine = currentLine.Replace(removeText, newText);
                                }

                                if (!currentLine.Contains(" PAYROLL REGISTER") 
                                 && !currentLine.Contains(" BANKERS FINANCIAL")
                                 && !currentLine.Contains("ENTITY")
                                 && !currentLine.Contains("EMPLOYEE")
                                 && !currentLine.Contains("CONTRIBUT")
                                 && currentLine != ""
                                 )
                                    payslipstring.AppendLine(currentLine);

                            }

                            payslipstring.AppendLine(currentLine);
                            //We now have a completed record
                            OutputSlip(payslipstring.ToString(),empName, headerLine.ToString());

                            payslipstring.Clear();
                        }

                    }
                }


            }

            
        }
    
        static void ExpandW2()
        {
            StringBuilder textstuff = new StringBuilder();
            string mainFileLoc = "c:\\temp\\W2\\2014\\W2_B01BIC_20141231_legal.pdf";
            string currentName = string.Empty;
            string year = "2015";
            Dictionary<string, int> idnumbers = GetIds();

            using (var pdfReader = new PdfReader(mainFileLoc))
            {
                int startpage = 1, endpage = 0;
                string datestring = string.Empty;
                PdfDocument outputDocument = new PdfDocument();
                string outputName = string.Empty;
                string formattedName = string.Empty;
                string uniqueName = string.Empty;
                int reviewcount = 1;

                // Loop through each page of the document
                for (var page = 1; page <= pdfReader.NumberOfPages; page++)
                {
                    ITextExtractionStrategy strategy = new SimpleTextExtractionStrategy();

                    //Get the PDF Page
                    var currentText = PdfTextExtractor.GetTextFromPage(
                    pdfReader,
                    page,
                    strategy);

                    currentText =
                        Encoding.UTF8.GetString(Encoding.Convert(
                            Encoding.Default,
                            Encoding.UTF8,
                            Encoding.Default.GetBytes(currentText)));

                   //Name of the employee
                   currentName = GetW2Name(currentText);

                    //turn FIRST M LAST to Last, First M
                    string[] tempName = currentName.Split(' ');

                    if (tempName.Length > 2)
                        formattedName = tempName[2] + ", " + tempName[0] + " " + tempName[1];
                    else
                        formattedName = tempName[1] + ", " + tempName[0];

                   
                   //Create file name
                    outputName = string.Format("c:\\temp\\W2\\2014\\{0}__{1}.pdf", currentName, year).Replace(" ", "");
                   //Create individual PDFS
                    ExtractPages(mainFileLoc, outputName, page, page);

                    currentName = currentName.Replace(" ", "");


                    AddW2IImportLine(currentName, year, formattedName, idnumbers);

                    Console.WriteLine("File extracted! " + currentName);
                    CreateImportFile("c:\\temp\\W2\\bfcinput.csv");
                   

                }

            }
            Console.ReadLine();
        }

        enum EmployeeTypes
        {
            Active,
            Terminated,
            L
        }

        static StringBuilder importTemplate = new StringBuilder();

        static void Main(string[] args)
        {
            ExtractPayslips();
          //  ExpandW2();

            return;

            StringBuilder textstuff = new StringBuilder();
            string mainFileLoc = "c:\\temp\\reviews2014B.pdf";
            Dictionary<string, int> idnumbers = new Dictionary<string, int>();
            List<string> reviewNames = new List<string>();

            //Get the dictionary of Last, First M - Employeeid
            idnumbers = GetIds();
         
            using (var pdfReader = new PdfReader(mainFileLoc))
            {
                int startpage = 1, endpage = 0;
                string datestring = string.Empty;
                PdfDocument outputDocument = new PdfDocument();
                string outputName = string.Empty;
                string name = string.Empty;
                string uniqueName = string.Empty;
                int reviewcount = 1;

                // Loop through each page of the document
                for (var page = 1; page <= pdfReader.NumberOfPages; page++) 
                {
                    ITextExtractionStrategy strategy = new SimpleTextExtractionStrategy();

                    //Get the PDF Page
                    var currentText = PdfTextExtractor.GetTextFromPage(
                    pdfReader,
                    page,
                    strategy);

                    currentText =
                        Encoding.UTF8.GetString(Encoding.Convert(
                            Encoding.Default,
                            Encoding.UTF8,
                            Encoding.Default.GetBytes(currentText)));

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
                         uniqueName = string.Format("{0}__{1}", name, datestring); ;
                        reviewNames.Add(name);

                        outputName = string.Format("c:\\temp\\{0}__{1}.pdf", name, datestring);//.Replace(" ", "");
                    }

                    if (currentText.Contains("Instructions for the Reviewer") && endpage > startpage)
                    {
                        Console.WriteLine(string.Format("Found new review at page {0} - Extracting...{1}", page.ToString(),name));

                        //Output the PDF
                        ExtractPages(mainFileLoc, outputName, startpage, endpage);

                        //Add a line for the import file
                        AddImportLine(uniqueName, datestring, name, idnumbers);
                        Console.WriteLine("File extracted!");

                        //Start the next file
                        startpage = endpage + 1;

                        //Get next date string
                        datestring = GetDatePos(currentText);

                        int startidx = currentText.LastIndexOf("Review") + 9;
                        int endidx = currentText.IndexOf("Page ", startidx) - 1;
                        int nameLength = endidx - startidx;

                        name = currentText.Substring(startidx, nameLength);

                        //Due to the fact that there might be 3 2/28/15 reviews, we need to ensure that each file is unique
                        uniqueName = string.Format("{0}__{1}", name, datestring);

                        while (reviewNames.Contains(uniqueName))
                        {
                            reviewcount++;
                            uniqueName = string.Format("{0}__{1}_{2}", name, datestring, reviewcount);
                        }

                        reviewcount = 1;
                        reviewNames.Add(uniqueName);

                        outputName = string.Format("c:\\temp\\{0}.pdf", uniqueName );
                      //outputName = string.Format("c:\\temp\\{0}__{1}.pdf", name, datestring).Replace(" ", "");
                    }

               
                    endpage++;

                }

            }

            CreateImportFile("c:\\temp\\employees\\bfcinput.csv");
            Console.ReadLine();
        }

        
        static Dictionary<string, int> GetIds()
        {
            Dictionary<string, int> theIds = new Dictionary<string, int>();
            string empFileLocA = "c:\\temp\\employees\\Active_Employees.csv";
            string empFileLocT = "c:\\temp\\employees\\Term_Employees.csv";
            string empFileLocL = "c:\\temp\\employees\\L_Employees.csv";

            /* Add active Employees to list */
            TextReader empFile = new StreamReader(empFileLocA);
            string line = string.Empty;

            while (empFile.Peek() != -1 )
            {
                line = empFile.ReadLine();
                string[] parts = line.Split(',');
                theIds.Add(string.Format("{0}, {1}", parts[0].Replace('"', ' ').Trim().ToUpper(), parts[1].Replace('"', ' ').Trim().ToUpper()), int.Parse(parts[2].Replace('"', ' ').Trim()));
            }

            empFile.Close();
            /*  We've gotten the active employees added */

            /* Add the termed employees */
            empFile = new StreamReader(empFileLocT);

            while (empFile.Peek() != -1)
            {
                line = empFile.ReadLine();
                string[] parts = line.Split(',');
                string key = string.Empty;
                int idvalue = -1;

                if (parts.Length == 4)
                {
                    //SPECIAL CASE
                    //The user must be Edenfield, IV, Edward James
                    key = "Edenfield, IV, Edward James";
                    idvalue = int.Parse(parts[3].Replace('"', ' ').Trim());
                }
                else
                {
                    //The users name Last, First M
                    key = string.Format("{0}, {1}", parts[0].Replace('"', ' ').Trim().ToUpper(), parts[1].Replace('"', ' ').Trim().ToUpper());
                    idvalue = int.Parse(parts[2].Replace('"', ' ').Trim());
                }
               
                if (!theIds.ContainsKey(key))
                {
                    theIds.Add(key,idvalue );
                }
            }

            empFile.Close();
            /* Termed Employees added to list */

            /* Added L Employeeds to List */
            empFile = new StreamReader(empFileLocL);

            while (empFile.Peek() != -1)
            {
                line = empFile.ReadLine();
                string[] parts = line.Split(',');
                //The users name Last, First M
                string key = string.Format("{0}, {1}", parts[0].Replace('"', ' ').Trim().ToUpper(), parts[1].Replace('"', ' ').Trim().ToUpper());
                int idvalue = int.Parse(parts[2].Replace('"', ' ').Trim());

                if (!theIds.ContainsKey(key))
                {
                    theIds.Add(key, idvalue);
                }
            }
            empFile.Close();
            /* L Employeeds have been added to List */


            return theIds; 

        }

        static string GetW2Name(string currentText)
        {
            string theName = string.Empty;
            int startIdx = currentText.IndexOf("ZIP Code") + 8;
            //Find Employee Box
            int employeeIdx = currentText.IndexOf("Employee's name, address, and ZIP code");
            int nameStartIdx = currentText.IndexOf('\n', employeeIdx) + 1;
            int nameEndIdx = currentText.IndexOf('\n', nameStartIdx);

            int nameLength = nameEndIdx - nameStartIdx;

            theName = currentText.Substring(nameStartIdx, nameLength);

            return theName; 
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

            //Employee's name, address, and ZIP code
        }

        static void AddW2IImportLine(string filename, string year, string name, Dictionary<string, int> idnumbers)
        {
            string importLine = string.Empty;
            string docName = string.Empty;
            string title = "W2 " + year;
            int employeeId = 0;

            employeeId = idnumbers.ContainsKey(name) ? idnumbers[name] : -1;

            string[] fileParts = filename.Split('_');

            docName = string.Format("{0}_{1}", filename, year);

        
            importLine = string.Format("{0},{1},{2},{3},{4},{5},{6},{7},{8}",
                                        "Client Data",
                                        employeeId.ToString(),
                                        "W2",
                                        docName,
                                        docName,
                                        "FILE",
                                        string.Format("\"{0}\"", docName),
                                        title,
                                        string.Format("\"{0}.pdf\"", filename));

            importTemplate.AppendLine(importLine);
        }

        static void AddImportLine(string filename, string period, string name, Dictionary<string, int> idnumbers)
        {
            string importLine = string.Empty;
            int employeeId = 0;
            string docName = string.Empty;
            string title = "Peformance Review" + period;

            employeeId = idnumbers.ContainsKey(name) ? idnumbers[name] : -1;

            string[] fileParts = filename.Split('_');

            docName = string.Format("Perf_{0}_{1}",employeeId,period);

            if (fileParts.Length == 6)
            {
                docName += "_" + fileParts[5];
                title += "_" + fileParts[5];
            }
             
            importLine = string.Format("{0},{1},{2},{3},{4},{5},{6},{7},{8}",
                                        "Client Data",
                                        employeeId.ToString(),
                                        "Performance Based",
                                        docName,
                                        docName,
                                        "FILE",
                                        string.Format("\"{0}\"",filename),
                                        title,
                                        string.Format("\"{0}.pdf\"", filename));

            importTemplate.AppendLine(importLine);
 
        }

        static void CreateImportFile(string filePath )
        {
            //"c:\\temp\\employees\\bfcinput.csv"
            TextWriter importFile = new StreamWriter(filePath);
            importFile.Write(importTemplate.ToString());
            importFile.Close();
        }

        //Output the specified page range to a Separate PDF
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
