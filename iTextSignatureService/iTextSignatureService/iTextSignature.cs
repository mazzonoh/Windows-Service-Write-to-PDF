using System;
using System.ServiceProcess;
using System.IO;

using iText.Kernel.Pdf;
using iText.Layout.Element;
using iText.Kernel.Font;
using iText.IO.Font.Constants;
using iText.Layout;
using iText.Layout.Properties;

namespace iTextSignatureService
{
    public partial class iTextSignature : ServiceBase
    {
        public iTextSignature()
        {
            InitializeComponent();
        }
       
        public const string jobsSourcePath = "D:\\JOBS"; //Folder to write the jobs for this service
        public const string documentSourcePath = "D:\\\\BASE_FILE.pdf"; //The PDF file to be written
        public const string documentDestinationPath = "D:\\DESTINATION"; //An output folder for the PDF after text changes
        public static string logFolder = Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles) + "\\iTextSignature"; //Log folder, usually the same where the service is installed

        /// <summary>
        /// To deploy this service, first build the solution with the appropriate template (debug or release) copy the contents of the bin\Debug or bin\Release folder of your project to the
        /// desired service installation folder (example: C:\Program Files (x86)\iTextSignature) then follow these steps:
        /// 
        /// Windows 8.1 and later releases installation instructions:
        ///	
        ///	Open CMD prompt as administrator
        ///	
        ///	Change directory to:
        ///	C:\Windows\Microsoft.NET\Framework\v4.0.30319
        ///	C:\Windows\Microsoft.NET\Framework64\v4.0.30319
        ///	
        ///	Install:
        ///	installutil.exe "C:\Program Files (x86)\iTextSignature\iTextSignature.exe"
        ///	
        ///	Uninstall: (if you are removing the service)
        ///	installutil /u "C:\Program Files (x86)\iTextSignature\iTextSignature.exe"
        ///	
        ///	Now type your username with domain (if necessary) and password, you can later change to a Local Account if you have any password expiration policy
        /// </summary>
        /// <param name="args"></param>
        protected override void OnStart(string[] args)
        {
            WriteLogLine("Service Started");
            WriteLogLine("Current configuration:");
            WriteLogLine("Jobs source path: " + jobsSourcePath);
            WriteLogLine("Blank document source path: " + documentSourcePath);
            WriteLogLine("Signed document destination path: " + documentDestinationPath);
            WriteLogLine("Log file: " + logFolder + "\\log.txt");

            //Debug interruption
            #if DEBUG
                WriteLogLine("DEBUG is ON");
                System.Diagnostics.Debugger.Launch();
            #endif
        }

        protected override void OnStop()
        {
            WriteLogLine("Service Stopped");
        }

        /// <summary>
        /// To send a custom command to this service, in C# use the following code from the application which are consuming this service:
        /// 
        /// //Call service to write PDF file
        /// ServiceController sc = new ServiceController("iTextSignature", Environment.MachineName);
        /// //this will grant permission to access the Service
        /// ServiceControllerPermission scp = new ServiceControllerPermission(ServiceControllerPermissionAccess.Control, Environment.MachineName, "iTextSignature");
        /// scp.Assert();
        /// sc.Refresh();
        /// //You can send any range from 128 to 255, be sure to check if there are running jobs before assigning a number, for example, if you check the jobsSourcePath
        /// //and find the job file 128.txt, be sure to check & send 129.txt and so on
        /// sc.ExecuteCommand(128); 
        /// 
        /// The job file must a txt with the following syntax:
        /// FileNameWithoutExtension|TextContent
        /// Example:
        /// My PDF|This is the new text
        /// 
        /// </summary>
        /// <param name="command"></param>
        protected override void OnCustomCommand(int command)
        {
            WriteLogLine("Command push: " + command);
            string filename = string.Empty;
            string textContent = string.Empty;

            WriteLogLine("Now opening " + jobsSourcePath + "\\" + command.ToString() + ".txt");
            using (StreamReader sReader = new StreamReader(jobsSourcePath + "\\" + command.ToString() + ".txt"))
            {
                while (sReader.Peek() >= 0)
                {
                    string line = sReader.ReadLine();
                    WriteLogLine("Processing line \"" + line + "\"");
                    filename = line.Split('|')[0];
                    textContent = line.Split('|')[1];
                }
            }

            if (WriteSignature(documentSourcePath, documentDestinationPath, filename, textContent) == true)
            {
                WriteLogLine("Text written with success for: " + filename);

                try
                {
                    //After completion, delete the job to free the allocated slot (128-255)
                    File.Delete(jobsSourcePath + "\\" + command.ToString() + ".txt");
                    WriteLogLine("File deleted: " + jobsSourcePath + "\\" + command.ToString() + ".txt");
                }

                catch
                {
                    WriteLogLine("Delete failed: " + jobsSourcePath + "\\" + command.ToString() + ".txt");
                }
            }

            else
            {
                WriteLogLine("Failed to write text for: " + filename);
            }
        }

        public static bool WriteSignature(string baseFile_ServerLocalPath, string destinationFile_serverLocalPath, string fileName, string textContent)
        {
            try
            {
                //Note that you acknowledge and agree with the AGPL license requirements when you call this method
                //Read more at: https://itextpdf.com/how-buy/AGPLv3-license
                //ATTENTION: THE LINE BELOW SHOULD NOT BE UNCOMMENTED WITHOUT FOLLOWING THE AGPL LICENSING AS EXPLAINED IN THE DEVELOPER'S PAGE ABOVE
                //iText.Commons.Actions.EventManager.AcknowledgeAgplUsageDisableWarningMessage();

                //Build file name and set input/output folders
                destinationFile_serverLocalPath += "\\" + fileName + ".pdf";
                PdfDocument pdfDoc = new PdfDocument(new PdfReader(baseFile_ServerLocalPath), new PdfWriter(destinationFile_serverLocalPath));

                //Repeat 'document.Add(new AreaBreak(AreaBreakType.NEXT_PAGE));' to move the working scope area across the pages, in this case the 5th page
                Document document = new Document(pdfDoc);
                document.Add(new AreaBreak(AreaBreakType.NEXT_PAGE));
                document.Add(new AreaBreak(AreaBreakType.NEXT_PAGE));
                document.Add(new AreaBreak(AreaBreakType.NEXT_PAGE));
                document.Add(new AreaBreak(AreaBreakType.NEXT_PAGE));

                PdfFont font = PdfFontFactory.CreateFont(StandardFonts.TIMES_BOLD);

                //Write text to specific position in the current page with custom formatting
                Text text = new Text(textContent).SetFont(font).SetFontSize(8);
                Paragraph pText = new Paragraph().Add(text).SetTextAlignment(TextAlignment.CENTER);                
                pText.SetFixedPosition(45, 240, 300);
                document.Add(pText);

                document.Close();
                pdfDoc.Close();

                return true;
            }

            catch
            {
                return false;
            }
        }

        /// <summary>
        /// This function will write logs in the service installation folder
        /// </summary>
        /// <param name="message"></param>
        public static void WriteLogLine(string message)
        {
            if (!Directory.Exists(logFolder)) { Directory.CreateDirectory(logFolder); }

            using (FileStream fs = new FileStream(logFolder + "\\log.txt", FileMode.Append, FileAccess.Write))
            {
                using (StreamWriter sw = new StreamWriter(fs))
                {
                    sw.WriteLine("[" + DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss") + "]: " + message);
                }
            }
        }
    }
}
