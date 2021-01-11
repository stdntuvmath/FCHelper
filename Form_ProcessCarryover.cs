using System;
using System.Windows.Forms;
using System.IO;
using System.Diagnostics;

namespace FCHelper_v001
{
    public partial class Form_ProcessCarryover : Form
    {


        private static string GroupsFolderPath = @"C:\Users\14025\Documents\File Consultants\Groups\";
        private static string GroupName;
        private static string TempFolderPath = @"C:\Users\14025\Documents\File Consultants\Brandon\Temp\";
        //private static string CarryoverFolder = @"C:\Users\14025\Documents\File Consultants\Groups\Payment Adjustment\Import\";
        private static string CarryoverFilePathAndName;
        private static string CarryoverFileNameOnly;
        private static string ERID;
        private static string EmailFilePathAndName;
        private static string EmailFileNameOnly;



        public Form_ProcessCarryover()
        {
            InitializeComponent();
        }

        private void Form_ProcessCarryover_Load(object sender, EventArgs e)
        {
            this.Location = new System.Drawing.Point(755, 230);
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
                    if (droppedFile.Contains(".txt"))
                    {

                        CarryoverFilePathAndName = droppedFile;
                        GetStringBetweenString getString = new GetStringBetweenString();



                        //MessageBox.Show(nonstandardCurrentFilePath);
                        try
                        {
                            CarryoverFileNameOnly = Path.GetFileName(CarryoverFilePathAndName);




                            listView1.Items.Add(CarryoverFileNameOnly);



                            //copy to Temp folder

                            string sourceFileandPath = CarryoverFilePathAndName;

                            string destinationFileandPath = TempFolderPath + CarryoverFileNameOnly;

                            if (!File.Exists(destinationFileandPath))
                            {
                                File.Copy(CarryoverFilePathAndName, destinationFileandPath);

                            }

                            //rename destination file
                            string destinationFile = TempFolderPath + CarryoverFileNameOnly;

                            string destinationFileNameOnly = Path.GetFileName(destinationFile);
                            string todaysDate = DateTime.Today.ToString("yyyyMMdd");

                            string newFileNameWithDate = todaysDate + "_" + CarryoverFileNameOnly;



                            RenameFile rename = new RenameFile();
                            rename.RenameFileMethod(destinationFile, newFileNameWithDate);




                            //open DOCS folder and Temp folder and import folder

                            ERID = getString.GetStringBetweenStringMethod(CarryoverFileNameOnly, "PFM.", "_");

                            GetGroupName getGroupName = new GetGroupName();
                            getGroupName.GetGroupNameMethod(ERID);
                            GroupName = GetGroupName.GroupName;


                            Process.Start(GroupsFolderPath + GroupName + @"\DOCS");
                            Process.Start(GroupsFolderPath + GroupName + @"\Import");

                            Process.Start(TempFolderPath);


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

                                string newFileNameWithDate = todaysDate + "_" + CarryoverFileNameOnly;



                                RenameFile rename = new RenameFile();
                                rename.RenameFileMethod(destinationFile, newFileNameWithDate);

                                //make an extra copy of the Email file
                                //if (!File.Exists(newEmailFilePathAndNameWithDate))
                                //{
                                //    File.Copy(CarryoverFilePathAndName, newEmailFilePathAndNameWithDate);

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

        private void button2_Click(object sender, EventArgs e)//Cancel
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
    }
}
