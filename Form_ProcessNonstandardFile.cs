using System;
using System.IO;
using System.Diagnostics;
using System.Windows.Forms;
using System.Threading;

namespace FCHelper_v001
{

    public partial class Form_ProcessNonstandardFile : Form
    {

        private static string GroupsFolderPath = @"C:\Users\14025\Documents\File Consultants\Groups\";
        private static string GroupName;
        private static string TempFolderPath = @"C:\Users\14025\Documents\File Consultants\Brandon\Temp\";
        private static string NonstandardFormPathAndName;
        private static string NonstandardFormFileNameOnly;
        private static string NonstandardFilePathAndName;
        private static string ERID;



        public Form_ProcessNonstandardFile()
        {
            InitializeComponent();


        }

        private void Form_ProcessNonstandardFile_Load(object sender, EventArgs e)
        {
            this.Location = new System.Drawing.Point(755, 230);

        }


        private void button1_Click(object sender, EventArgs e)//OK TO PROCESS Button
        {
            listView1.Clear();
            File.Delete(TempFolderPath+ NonstandardFormFileNameOnly);
            
            Process.Start(GroupsFolderPath+GroupName+@"\Import");
            Process.Start(TempFolderPath);


            PopulateNonstandardWikiLog populate = new PopulateNonstandardWikiLog();
            populate.PopulateNonstandardWikiLogMethod();

            CreateTestResultOutlookEmail createEmail = new CreateTestResultOutlookEmail();

            foreach (string erid in GetNonstandardFileData.EmployerID)
            {
                ERID = erid;
            }


            GetStringBetweenString getStringBetweenString = new GetStringBetweenString();
            

            string requestersFirstName = getStringBetweenString.GetStringBetweenStringMethod(GetNonstandardFileData.Requester,""," ");
            string managersLastName = getStringBetweenString.GetStringBetweenStringMethod(GetNonstandardFileData.ApprovingManager, " ", "•");


            string emailTo = GetNonstandardFileData.Requester;
            string emailCC = GetNonstandardFileData.ApprovingManager;
            string emailSubject = GetNonstandardFileData.EmployerName+" - "+ ERID + " - Nonstandard File Request Approved";
            string emailBody = String.Format("<p style = \"font-size:11pt;\">Hello "+ requestersFirstName +
                                ",<br/><br/>" +
                "Your file(s) have been staged for processing and should " +
                "produce results within the next few hours.</p> ");

            createEmail.CreateTestResultOutlookEmailMethod(emailTo,string.Empty, string.Empty, string.Empty, 
                        emailCC, string.Empty, string.Empty, string.Empty, emailSubject, emailBody);
            
        }

        private void button2_Click(object sender, EventArgs e)//Cancel button
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
                    if (droppedFile.Contains(".doc") || droppedFile.Contains(".docx") || droppedFile.Contains("Nonstandard") || droppedFile.Contains("NonstandardFileProcessRequest"))
                    {


                        NonstandardFormPathAndName = droppedFile;
                        //MessageBox.Show(nonstandardCurrentFilePath);
                        try
                        {
                            string nonstandardFileNameOnly = Path.GetFileName(NonstandardFormPathAndName);
                            NonstandardFormFileNameOnly = nonstandardFileNameOnly;
                            listView1.Items.Add(nonstandardFileNameOnly);

                            string sourceFileandPath = NonstandardFormPathAndName;

                            string destinationFileandPath = TempFolderPath + nonstandardFileNameOnly;

                            if (!File.Exists(destinationFileandPath))
                            {
                                File.Copy(NonstandardFormPathAndName, destinationFileandPath);

                            }
                            else
                            {
                                MessageBox.Show("File Already Exists");
                            }

                        }
                        catch (UnauthorizedAccessException ex)
                        {

                        }

                        //pull data from table
                        GetNonstandardFileData getData = new GetNonstandardFileData();
                        getData.GetNonstandardFileDataMethod(droppedFile);

                        string[] ERIDs = GetNonstandardFileData.EmployerID;

                        foreach (string erid in ERIDs)
                        {

                            //MessageBox.Show("ERID going inside Method: "+erid);

                            GetGroupName getGroupName = new GetGroupName();
                            getGroupName.GetGroupNameMethod(erid);
                            string groupName = GetGroupName.GroupName;
                            GroupName = groupName;


                            //MessageBox.Show("groupName coming outside of method: "+groupName);

                            


                            //MessageBox.Show(GroupName);

                            //MessageBox.Show(GroupsFolderPath + GroupName);


                            if (Directory.Exists(GroupsFolderPath + GroupName + @"\DOCS"))
                            {

                                //MessageBox.Show("DOCS");
                                Process.Start(GroupsFolderPath + GroupName + @"\DOCS");
                                //Process.Start(@"https://etltrac.payflex.com/etl/");

                                //Thread.Sleep(6500);

                                //MOUSE_LeftClick leftClick = new MOUSE_LeftClick();
                                //leftClick.MOUSE_LeftClickMethod(3500, 145);

                                //SendKeys.Send(erid);
                                //SendKeys.Send("{ENTER}");
                            }
                            else if (Directory.Exists(GroupsFolderPath + GroupName + @"\Docs"))
                            {

                                //MessageBox.Show("Docs");

                                Process.Start(GroupsFolderPath + GroupName + @"\Docs");
                                //Process.Start(@"https://etltrac.payflex.com/etl/");

                                //Thread.Sleep(6500);

                                //MOUSE_LeftClick leftClick = new MOUSE_LeftClick();
                                //leftClick.MOUSE_LeftClickMethod(3500, 145);

                                //SendKeys.Send(erid);
                                //SendKeys.Send("{ENTER}");
                            }
                            else if (Directory.Exists(GroupsFolderPath + GroupName + @"\docs"))
                            {

                                //MessageBox.Show("docs");

                                //Process.Start(GroupsFolderPath + GroupName + @"\docs");
                                //Process.Start(@"https://etltrac.payflex.com/etl/");

                                //Thread.Sleep(3500);

                                //MOUSE_LeftClick leftClick = new MOUSE_LeftClick();
                                //leftClick.MOUSE_LeftClickMethod(3500, 145);

                                //SendKeys.Send(erid);
                                //SendKeys.Send("{ENTER}");
                            }


                        }

                        

                        //MessageBox.Show(groupName);

                        //open docs folder

                        


                    }
                    else if (droppedFile.Contains(".txt") || droppedFile.Contains(".pgp"))
                    {

                        NonstandardFilePathAndName = droppedFile;

                        string fileNameOnly = Path.GetFileName(NonstandardFilePathAndName);
                        string sourceFile = NonstandardFilePathAndName;
                        string destinationFile = TempFolderPath + fileNameOnly;
                        listView1.Items.Add(fileNameOnly);

                        if (!File.Exists(fileNameOnly))
                        {
                            try
                            {
                                File.Copy(sourceFile, destinationFile);
                                

                                //rename destination file

                                string destinationFileNameOnly = Path.GetFileName(destinationFile);
                                string todaysDate = DateTime.Today.ToString("yyyyMMdd");

                                string newFileNameWithDate = todaysDate + "_" + destinationFileNameOnly;
                                string newNSFormFilePathAndNameWithDate = TempFolderPath + newFileNameWithDate + ".docx";
                                RenameFile rename = new RenameFile();
                                rename.RenameFileMethod(destinationFile, newFileNameWithDate);

                                //make an extra copy of the nonstandard form file
                                if (!File.Exists(newNSFormFilePathAndNameWithDate))
                                {
                                    File.Copy(NonstandardFormPathAndName, newNSFormFilePathAndNameWithDate);

                                }
                            }
                            catch (UnauthorizedAccessException ex)
                            {

                            }
                        }

                    }
                    if (droppedFile.Contains(".msg"))
                    {
                        //MessageBox.Show("Email File " + droppedFile);

                        


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
