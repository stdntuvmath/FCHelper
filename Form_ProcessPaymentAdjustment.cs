using System;
using System.IO;
using System.Threading;
using System.Windows.Forms;
using System.Diagnostics;


namespace FCHelper_v001
{
    public partial class Form_ProcessPaymentAdjustment : Form
    {

        public static string EmailSubjectLine;



        private static string GroupsFolderPath = @"C:\Users\14025\Documents\File Consultants\Groups\";
        private static string GroupName;
        private static string TempFolderPath = @"C:\Users\14025\Documents\File Consultants\Brandon\Temp\";
        private static string PaymentAdjustmentFolder = @"C:\Users\14025\Documents\File Consultants\Groups\Payment Adjustment\Import\";
        private static string PaymentAdjustmentFilePathAndName;
        private static string PaymentAdjustmentFileNameOnly;
        private static string EmailFilePathAndName;
        private static string EmailFileNameOnly;
        private static string ERID;
        private static string CentralFolder = "PayFlex File Support Central";
        private static string RootFolder = "TurnerB1@aetna.com";

        GetStringBetweenString getString = new GetStringBetweenString();




        public Form_ProcessPaymentAdjustment()
        {
            InitializeComponent();
        }

        private void Form_ProcessPaymentAdjustment_Load(object sender, EventArgs e)
        {
            this.Location = new System.Drawing.Point(755, 230);

        }
        private void button1_Click(object sender, EventArgs e)//OK TO PROCESS Button
        {
            string emailFileNameWithoutExtension = getString.GetStringBetweenStringMethod(EmailFileNameOnly,"",".msg");

           //GetOutlookEmailInformation getInfo = new GetOutlookEmailInformation();
            //getInfo.GetOutlookEmailInformationMethod(emailFileNameWithoutExtension, RootFolder);
        }//OK TO PROCESS Button

        private void button2_Click(object sender, EventArgs e)//CANCEL Button
        {
            //delete temp files

            string directory = TempFolderPath;


            if (Directory.Exists(directory))
            {
                System.IO.DirectoryInfo di = new DirectoryInfo(directory);

                foreach (FileInfo file in di.GetFiles())
                {
                    file.Delete();
                }


            }

            this.Dispose();
            this.Close();
        }

        private void listView1_DragEnter(object sender, DragEventArgs e)
        {
            e.Effect = DragDropEffects.All;
        }

        private void listView1_DragDrop(object sender, DragEventArgs e)
        {
            string[] droppedFiles = (string[])e.Data.GetData(DataFormats.FileDrop, false);
            try
            {
                foreach (string droppedFile in droppedFiles)//iterate through all files dropped into the form
                {
                    if (droppedFile.Contains(".dat"))
                    {

                        Process.Start(@"C:\Users\14025\Documents\File Consultants\Groups\Payment Adjustment\Import\");
                        PaymentAdjustmentFilePathAndName = droppedFile;
                        GetStringBetweenString getString = new GetStringBetweenString();



                        //MessageBox.Show(nonstandardCurrentFilePath);
                        try
                        {
                            PaymentAdjustmentFileNameOnly = Path.GetFileName(PaymentAdjustmentFilePathAndName);
                        
                            listView1.Items.Add(PaymentAdjustmentFileNameOnly);

                                                       

                            //copy to Temp folder

                            string sourceFileandPath = PaymentAdjustmentFilePathAndName;

                            string destinationFileandPath = TempFolderPath + PaymentAdjustmentFileNameOnly;

                            if (!File.Exists(destinationFileandPath))
                            {
                                File.Copy(PaymentAdjustmentFilePathAndName, destinationFileandPath);

                            }


                            //copy to Payment Adjustment folder

                            //destinationFileandPath = PaymentAdjustmentFolder + PaymentAdjustmentFileNameOnly;

                            //if (!File.Exists(destinationFileandPath))
                            //{
                            //    File.Copy(PaymentAdjustmentFilePathAndName, destinationFileandPath);

                            //}

                            //open DOCS folder and Temp folder

                            ERID = getString.GetStringBetweenStringMethod(PaymentAdjustmentFileNameOnly, "", "_");

                            GetGroupName getGroupName = new GetGroupName();
                            getGroupName.GetGroupNameMethod(ERID);
                            GroupName = GetGroupName.GroupName;


                            Process.Start(GroupsFolderPath+GroupName+@"\DOCS");
                            Process.Start(TempFolderPath);

                            //save filename in clipboard for copy/paste into script
                            Clipboard.SetText(PaymentAdjustmentFileNameOnly);

                        }
                        catch (UnauthorizedAccessException ex)
                        {

                        }

                        


                    }
                    else if (droppedFile.Contains(".msg"))
                    {

                        EmailFilePathAndName = droppedFile;

                        EmailFileNameOnly = Path.GetFileName(EmailFilePathAndName);

                        listView1.Items.Add(EmailFileNameOnly);

                        //copy to temp folder with timestamp and pmtadj filename
                        string sourceFile = EmailFilePathAndName;
                        string destinationFile = TempFolderPath + EmailFileNameOnly;

                        if (!File.Exists(destinationFile))
                        {
                            try
                            {
                                File.Copy(sourceFile, destinationFile);


                                //rename destination file

                                string destinationFileNameOnly = Path.GetFileName(destinationFile);
                                string todaysDate = DateTime.Today.ToString("yyyyMMdd");

                                string newFileNameWithDate = todaysDate + "_" + PaymentAdjustmentFileNameOnly;



                                RenameFile rename = new RenameFile();
                                rename.RenameFileMethod(destinationFile, newFileNameWithDate);

                                //make an extra copy of the Email file
                                //if (!File.Exists(newEmailFilePathAndNameWithDate))
                                //{
                                //    File.Copy(PaymentAdjustmentFilePathAndName, newEmailFilePathAndNameWithDate);

                                //}
                            }
                            catch (UnauthorizedAccessException ex)
                            {

                            }
                        }

                    }
                    

                    this.BringToFront();
                }
            }
            catch (ArgumentNullException ex)
            {

            }
        }


    }
}
