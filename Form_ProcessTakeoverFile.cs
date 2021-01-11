using System;
using System.IO;
using System.Threading;
using System.Windows.Forms;
using System.Diagnostics;

namespace FCHelper_v001
{
    public partial class Form_ProcessTakeoverFile : Form
    {

        private static string TakeoverFilePathAndName;
        private static string TakeoverFileNameOnly;
        private static string ImportedTakeoverFilePathAndName;
        private static string ERID;
        private static string TempFolderPath = @"C:\Users\14025\Documents\File Consultants\Brandon\Temp\";
        private static string GroupsFolderPath = @"C:\Users\14025\Documents\File Consultants\Groups\";
        private static string GroupName;
        private static string TakeoverDatabaseOriginalPath = @"\\phx-fs-02.payflex.com\GDrive\Data\PFS\GFP\group\~DOCS\! NEW DATABASE TEMPLATE\Takeover Template\CBAS_Takeover_v4.Meow.accdb";
        private static string TakeoverDatabaseFilePath;
        private static string TakeoverDatabaseFileNameOnly = "CBAS_Takeover_v4.Meow.accdb";
        private static string EmailFilePathAndName;
        private static string EmailFileNameOnly;
        private static string InputPSVFilePathAndName;
        private static string InputPSVFileNameOnly;

        public Form_ProcessTakeoverFile()
        {
            InitializeComponent();
        }

        private void Form_ProcessTakeoverFile_Load(object sender, EventArgs e)
        {
            this.Location = new System.Drawing.Point(755, 230);
        }

        private void listView1_DragEnter(object sender, DragEventArgs e)
        {
            e.Effect = DragDropEffects.All;

        }

        private void listView1_DragDrop(object sender, DragEventArgs e)
        {
            GetStringBetweenString getString = new GetStringBetweenString();


            string[] droppedFiles = (string[])e.Data.GetData(DataFormats.FileDrop, false);
            try
            {
                foreach (string droppedFile in droppedFiles)//iterate through all files dropped into the form
                {
                    if (droppedFile.Contains("_Takeover_"))
                    {

                        TakeoverFilePathAndName = droppedFile;






                        //MessageBox.Show(nonstandardCurrentFilePath);
                        try
                        {
                            TakeoverFileNameOnly = Path.GetFileName(TakeoverFilePathAndName);

                            //get ERID from filename

                            ERID = getString.GetStringBetweenStringMethod(TakeoverFileNameOnly, "", "_");

                            listView1.Items.Add(TakeoverFileNameOnly);

                            //copy Takeover database to employer main folder

                            GetGroupName getName = new GetGroupName();
                            getName.GetGroupNameMethod(ERID);

                            GroupName = GetGroupName.GroupName;


                            string sourceFileandPath = TakeoverDatabaseOriginalPath;

                            string destinationFileandPath = GroupsFolderPath + GroupName + @"\"+TakeoverDatabaseFileNameOnly;

                            TakeoverDatabaseFilePath = destinationFileandPath;

                            if (!File.Exists(destinationFileandPath))
                            {
                                try
                                {
                                    File.Copy(sourceFileandPath, destinationFileandPath);
                                }
                                catch (IOException ex)
                                {
                                    MessageBox.Show("The file " + TakeoverDatabaseOriginalPath + " cannot be copied to the groups main folder because: \r\r", "I/O Exception", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                }

                            }

                            //copy the takeover file to the groups import folder


                            sourceFileandPath = TakeoverFilePathAndName;

                            destinationFileandPath = GroupsFolderPath + GroupName + @"\Import\" + TakeoverFileNameOnly;

                            ImportedTakeoverFilePathAndName = destinationFileandPath;

                            if (!File.Exists(destinationFileandPath))
                            {
                                try
                                {
                                    File.Copy(sourceFileandPath, destinationFileandPath);
                                }
                                catch (IOException ex)
                                {
                                    MessageBox.Show("The file " + TakeoverFilePathAndName + " cannot be copied to the groups import folder because: \r\r", "I/O Exception", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                }

                            }


                            

                            //open all folders and Temp folder                        
                            Process.Start(GroupsFolderPath + GroupName);
                            //Process.Start(GroupsFolderPath + GroupName + @"\Import");
                            Process.Start(GroupsFolderPath + GroupName + @"\DOCS");
                            Process.Start(TempFolderPath);

                            //save file path to clipboard

                            Clipboard.SetText(GroupsFolderPath + GroupName + @"\Import");


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

                                string newFileNameWithDate = todaysDate + "_" + TakeoverFileNameOnly;



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
                    else if (droppedFile.Contains(".psv"))
                    {
                        InputPSVFilePathAndName = droppedFile;

                        //MessageBox.Show(InputPSVFilePathAndName);


                        InputPSVFileNameOnly = Path.GetFileName(InputPSVFilePathAndName);

                        //MessageBox.Show(InputPSVFileNameOnly);


                        //get processed datetime stamp

                        string dateTimeStampWithGroupName = getString.GetStringBetweenStringMethod(InputPSVFileNameOnly,"","-");

                        //MessageBox.Show(dateTimeStampWithGroupName);

                        //get first letter of the group name

                        string substring1 = dateTimeStampWithGroupName.Substring(dateTimeStampWithGroupName.IndexOf('_'));

                        //MessageBox.Show(substring1);

                        string substring2 = substring1.Substring(1, substring1.Length-1);

                        //MessageBox.Show(substring2);

                        string substring3 = substring2.Substring(substring2.IndexOf('_'));

                        //MessageBox.Show(substring3);


                        string substring4 = substring3.Substring(1, substring3.Length - 1);

                        //MessageBox.Show(substring4);

                        string firstLetter = substring4.Substring(0,1);


                        //MessageBox.Show(firstLetter);


                        string dateTimeStamp = getString.GetStringBetweenStringMethod(dateTimeStampWithGroupName,"",firstLetter);

                        //MessageBox.Show(dateTimeStamp);


                        //rename takeover database file


                        string newDatabaseFileNameWithstamp = dateTimeStamp + TakeoverDatabaseFileNameOnly;



                        RenameFile rename = new RenameFile();
                        rename.RenameFileMethod(TakeoverDatabaseFilePath, newDatabaseFileNameWithstamp);


                        //rename takeover file


                        string newTakeoverFileNameWithstamp = dateTimeStamp + TakeoverFileNameOnly;



                        
                        rename.RenameFileMethod(ImportedTakeoverFilePathAndName, newTakeoverFileNameWithstamp);


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
