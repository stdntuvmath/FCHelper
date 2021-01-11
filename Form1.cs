using System;
using System.ComponentModel;
using System.Linq;
using System.Windows.Forms;
using System.Diagnostics;
using System.IO;

namespace FCHelper_v001
{
    /*
     
        20210110 - v1.001 - BT - Updated icon
         
         
         
         */
    


    public partial class Form_Main : Form
    {

       
        //Global Variables
        public static string Fname;
        public static string ERID;
        public static string ClientID;
        public static string OrigFileNameWithOutPFPFM;
        public static string GroupName;
        public static string GroupsFolderPath = @"C:\Users\14025\Documents\File Consultants\Groups\";
        public static bool TextBox1Enter = false;
        private static bool NoDateTimeStampsGiven = false;






        public Form_Main()
        {
            InitializeComponent();

            //Control tab indexes
            textBox1.TabIndex = 0;
            button2.TabIndex = 1;
            button1.TabIndex = 2;
            textBox3.TabIndex = 3;
            textBox5.TabIndex = 4;
            
        }

        private void Form_Main_Load(object sender, EventArgs e)
        {
            /*Both
             
            UpdateFCImpTracker update = new UpdateFCImpTracker();
            update.UpdateFCImpTrackerMethod();
             
            and

            ETLAutomatedEmails monitorEmails = new ETLAutomatedEmails();
            monitorEmails.ETLAutomatedEmailsMethod();

            cause a Symantec antivirus popup and cause the cyber security department to
            by notified. Unconfirmed but it seems to occur when each method tries to search
            for the ERID in my ImpList spreadsheet.

             */

            //UpdateFCImpTracker update = new UpdateFCImpTracker();
            //update.UpdateFCImpTrackerMethod();

            this.Location = new System.Drawing.Point(0, 230);
            timer1.Start();

            try
            {
                backgroundWorker1.RunWorkerAsync(1000);//ETL email automation activation

            }
            catch (InvalidOperationException ex)
            {
                MessageBox.Show("The exception that is thrown when a method call is invalid for the object's current state.\r\r" + ex, "ETLAutomatedEmails Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }


        }

        private void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
        {
            ETLAutomatedEmails monitorEmails = new ETLAutomatedEmails();
            //monitorEmails.ETLAutomatedEmailsMethod();
        }


        public static string getStringBetweenString(string givenString, string stringYouWantBegin, string stringYouWantEnd)
        {
            int Start, End;


            if (givenString.Contains(stringYouWantBegin) && givenString.Contains(stringYouWantEnd))
            {
                Start = givenString.IndexOf(stringYouWantBegin, 0) + stringYouWantBegin.Length;
                End = givenString.IndexOf(stringYouWantEnd, Start);

                return givenString.Substring(Start, End - Start);
            }
            else
            {
                return "";
            }
        }



        private void openFileToolStripMenuItem_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            openFileDialog1.DefaultExt = "*.*";
            openFileDialog1.InitialDirectory = @"C:\Users\14025\Documents\File Consultants\Groups";

            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                string fileNameAndPath = openFileDialog1.FileName;
                string fileName = Path.GetFileName(fileNameAndPath);
                textBox3.Text = fileName;
                if (fileName.Contains("PF.PFM"))
                {
                    string origFileName = fileName.Substring(fileName.IndexOf('P'));
                    textBox5.Text = origFileName;

                    string erid = getStringBetweenString(origFileName, "M.", "_");

                    textBox1.Text = erid;
                }
                else
                {
                    textBox5.Text = "";
                }

                Stream myStream;
                if ((myStream = openFileDialog1.OpenFile()) != null)
                {
                    string fileText = File.ReadAllText(fileNameAndPath);
                    richTextBox1.Text = fileText;
                }//if we need to read the file and show it in a box
            }
        }

        private void openGroupsFolderToolStripMenuItem1_Click(object sender, EventArgs e)
        {



        }




        private void AddImplementationToolStripMenuItem_Click(object sender, EventArgs e)
        {
            
            Form_AddImplementation form = new Form_AddImplementation();
            form.Show();
        }

        private void openImplementationToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Form_OpenImplementationList form = new Form_OpenImplementationList();
            form.Show();
        }





        private void claimsTroubleshootingToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Form_ClaimsTroubleshooting form = new Form_ClaimsTroubleshooting();
            form.Show();
        }

        private void createConnectedClaimsFolderToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //prompt user for ERID

            //string input = Microsoft.VisualBasic.Interaction.InputBox("Please enter the employer's ID","Enter Employer ID","",400,300);
            Form_Connected_Claims_Folder_Prompt form = new Form_Connected_Claims_Folder_Prompt();
            form.Show();

            

        }




        private void nonstandardToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Form_ProcessNonstandardFile form_ProcessNonstandardFile = new Form_ProcessNonstandardFile();
            form_ProcessNonstandardFile.Show();
        }

        private void paymentAdjustmentToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Form_ProcessPaymentAdjustment form_ProcessPaymentAdjustment = new Form_ProcessPaymentAdjustment();
            form_ProcessPaymentAdjustment.Show();
        }

        private void carryoverFileToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Form_ProcessCarryover formCarryover = new Form_ProcessCarryover();
            formCarryover.Show();
        }

        private void takeoverFileToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Form_ProcessTakeoverFile processTakeoverFile = new Form_ProcessTakeoverFile();
            processTakeoverFile.Show();
        }



        private void timer1_Tick(object sender, EventArgs e)
        {
            label3.Text = DateTime.Now.ToString("T");
        }

        private void button1_Click_2(object sender, EventArgs e)//clear all controls
        {

            ClientID = null;
            ERID = null;

            foreach (Control control in this.Controls)
            {
                if (control is TextBox)
                {
                    TextBox txtbox = (TextBox)control;
                    txtbox.Text = string.Empty;
                }
                else if (control is RichTextBox)
                {
                    RichTextBox richTextBox = (RichTextBox)control;
                    richTextBox.Text = string.Empty;
                }
                else if (control is ListView)
                {
                    ListView list = (ListView)control;
                    list.Items.Clear();
                }
                
            }

            this.Text = "FC Helper";
            //delete temp files

            string directory = @"C:\Users\14025\Documents\File Consultants\Brandon\Temp\";


            if (Directory.Exists(directory))
            {
                System.IO.DirectoryInfo di = new DirectoryInfo(directory);

                foreach (FileInfo file in di.GetFiles())
                {
                    file.Delete();
                }
                

            }

        }

        private void button2_Click(object sender, EventArgs e)//Open Employer Folder
        {

            if (textBox1.Text == "")
            {

            }
            else
            {
                GetGroupName nameAndER = new GetGroupName();
                nameAndER.GetGroupNameMethod(textBox1.Text);
                string NameAndERID = GetGroupName.GroupName;

                Process.Start(@"C:\Users\14025\Documents\File Consultants\Groups\" + NameAndERID);
            }

            

        }

        private void button3_Click(object sender, EventArgs e)//Open Temp Folder
        {
            Process.Start(@"C:\Users\14025\Documents\File Consultants\Brandon\Temp");
        }


        private void Form_Main_DragDrop(object sender, DragEventArgs e)
        {
            //GetGroupName.GroupName = null;
            try
            {
                string[] droppedFileToOpen = (string[])e.Data.GetData(DataFormats.FileDrop);
                int counter = 0;


                //get dropped filename
                foreach (string droppedFilesCurrentPath in droppedFileToOpen)
                {

                    string fName = Path.GetFileName(droppedFilesCurrentPath);


                    int countUnderscores = fName.Count(x => x == '_');//counting the number of underscores in filename
                    int countPeriods = fName.Count(x => x == '.');//counting the number of periods in filename

                    int totalFilenameCharacterLength = fName.Length;//


                    Fname = fName;
                    textBox3.Text = fName;

                    
                        int firstLetterIndex = fName.IndexOf('P');






                    if (firstLetterIndex == 0 && fName.Contains("PF_TEST"))
                    {

                        string subStringTEST_ = Fname.Substring(Fname.IndexOf('T'));



                        string originalFileName = subStringTEST_.Substring(subStringTEST_.IndexOf('P'));





                        string erid = getStringBetweenString(originalFileName, "M.", "_");



                        textBox1.Text = erid;
                        ERID = erid;

                        //get group name (without carriers) from erid

                        GetGroupName getGroupName = new GetGroupName();
                        getGroupName.GetGroupNameMethod(ERID);
                        string groupName = GetGroupName.GroupName;

                        

                        this.Text = "FC Helper - "+ groupName;

                        //get client ID
                        //GetClientID getClientID = new GetClientID();
                        //getClientID.GetClientIDMethod(groupName);


                        //ClientID = GetClientID.ClientID;

                        //textBox2.Text = ClientID;









                        //get group folder file path
                        string groupFolderFilePath = @"C:\Users\14025\Documents\File Consultants\Groups\" + groupName;

                        try
                        {
                            string[] groupFolderFilePathOfFile = Directory.GetFiles(groupFolderFilePath, "*" + Fname + ".*", SearchOption.AllDirectories);

                            textBox6.Text = "NA";
                            textBox5.Text = originalFileName;
                            listView1.Items.Add("File has not processed yet.");

                            foreach (string file in groupFolderFilePathOfFile)
                            {


                                if (!File.Exists(file))
                                {

                                    textBox6.Text = "NA";
                                    textBox5.Text = originalFileName;
                                    listView1.Items.Add("File has not processed yet.");
                                    listView2.Items.Add("File has not processed yet.");
                                    listView3.Items.Add("File has not processed yet.");


                                }
                                else if (File.Exists(file))
                                {

                                    listView1.Clear();
                                    textBox6.Clear();

                                    string fileName = Path.GetFileName(file);
                                    Fname = null;
                                    Fname = fileName;
                                    textBox3.Text = fileName;
                                    //textBox5.Text = originalFileName;

                                    string processedDateTimeStampWithExcess1 = getStringBetweenString(fileName, "", "_2");
                                    string processedDateTimeStampWithExcess2 = getStringBetweenString(fileName, "", "_P");


                                    string dateStamp = getStringBetweenString(processedDateTimeStampWithExcess2, "", "_");
                                    string timeStamp = getStringBetweenString(processedDateTimeStampWithExcess2, "_", "_");

                                    string processedDateTimeStamp = dateStamp + "_" + timeStamp;


                                    textBox6.Text = processedDateTimeStamp;

                                    string destinationPath = @"C:\Users\14025\Documents\File Consultants\Brandon\Temp\" + fName;
                                    //read file in RichTextbox and Copy/Paste the file in question
                                    if (!File.Exists(destinationPath))
                                    {
                                        File.Copy(file, destinationPath);
                                        listView1.Items.Add(textBox5.Text);
                                        string fileText = File.ReadAllText(destinationPath);
                                        richTextBox1.Text = fileText;
                                    }
                                    else
                                    {
                                        MessageBox.Show("This file already exists in the Temp folder.", "Can't Do It. Wouldn't be Proper.", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                                    }



                                    //search for the groups error folder error reports
                                    GetErrorReportFileName name = new GetErrorReportFileName();
                                    name.GetErrorReportFileNameMethod(erid, processedDateTimeStamp);

                                    //search for the groups error folder for psvs
                                    GetInputPSVFileNames names = new GetInputPSVFileNames();
                                    names.GetInputPSVFileNamesMethod(erid, processedDateTimeStamp);

                                    //copy error reports to temp folder

                                    foreach (string groupFolderItem in GetErrorReportFileName.ErrorReportGroupFolderLocations)
                                    {

                                        //get source path+filenames
                                        string sourceFile = groupFolderItem;


                                        string errorReportfilename = Path.GetFileName(sourceFile);
                                        string destinationFile = @"C:\Users\14025\Documents\File Consultants\Brandon\Temp\" + errorReportfilename;

                                        if (!File.Exists(destinationFile))
                                        {
                                            File.Copy(sourceFile, destinationFile);

                                        }
                                    }

                                    //copy psv's to temp folder
                                    foreach (string groupFolderItem in GetInputPSVFileNames.PSVFolderLocations)
                                    {


                                        //get source path+filenames
                                        string sourceFile = groupFolderItem;


                                        string psvfilename = Path.GetFileName(sourceFile);
                                        string destinationFile = @"C:\Users\14025\Documents\File Consultants\Brandon\Temp\" + psvfilename;

                                        if (!File.Exists(destinationFile))
                                        {
                                            File.Copy(sourceFile, destinationFile);

                                        }



                                    }

                                    //add error report to listview2

                                    foreach (string file1 in GetErrorReportFileName.ErrorReportFileNamesOnly)
                                    {
                                        listView2.Items.Add(file1);

                                    }
                                    //add inputpsvs to listview3

                                    foreach (string file1 in GetInputPSVFileNames.InputPSVFileNamesOnly)
                                    {

                                        listView3.Items.Add(file1);

                                    }
                                    //add outputpsvs to listview4

                                    foreach (string file1 in GetInputPSVFileNames.OutputPSVFileNamesOnly)
                                    {
                                        //MessageBox.Show(file);

                                        listView4.Items.Add(file1);

                                    }

                                }
                            }
                        }
                        catch (ArgumentNullException ex)
                        {
                            MessageBox.Show("No file found.");
                        }


                    }
                    else if (!fName.Contains("PF.PFM") && !fName.Contains(".xlsm")&& !(firstLetterIndex == 0))
                    {
                        textBox5.Text = fName;

                        string dateProcessed = getStringBetweenString(fName, "", "_");
                        string timeProcessed = getStringBetweenString(fName, "_", "_");

                        string processedFileName = getStringBetweenString(fName,"","");
                        MessageBox.Show(processedFileName);


                        //int indexOfFirstUnderscore = fName.IndexOf('_');
                        //string withoutDateProcessed = fName.Substring(indexOfFirstUnderscore);
                        //MessageBox.Show(withoutDateProcessed);

                        //int indexOfSecondUnderscore = withoutDateProcessed.IndexOf('_');
                        //string withoutTimeProcessed = withoutDateProcessed.Substring(indexOfSecondUnderscore);
                        //MessageBox.Show(withoutTimeProcessed);

                        //int indexOfThirdUnderscore = withoutTimeProcessed.IndexOf('_');
                        //string withoutDateReceived = withoutTimeProcessed.Substring(indexOfThirdUnderscore);
                        //MessageBox.Show(withoutDateReceived);

                        //int indexOfFourthUnderscore = withoutDateReceived.IndexOf('_');
                        //string fileNameWithNoStamps = withoutDateReceived.Substring(indexOfFourthUnderscore);
                        //MessageBox.Show(fileNameWithNoStamps);

                        //string erid = getStringBetweenString(fileNameWithNoStamps, "", "_");

                        //textBox1.Text = erid;
                        //ERID = erid;
                        //MessageBox.Show(erid);

                        ////get group name (without carriers) from erid

                        //GetGroupName getGroupName = new GetGroupName();
                        //getGroupName.GetGroupNameMethod(ERID);
                        //string groupName = GetGroupName.GroupName;

                        //MessageBox.Show(groupName);
                    }

                    else if (firstLetterIndex == 0 && fName.Contains("PF.PFM"))
                    {

                        //Fname = fName;

                        //string origFileNameWithPF_TEST = fName.Substring(fName.IndexOf('P'));


                        //string subStringTEST_ = Fname.Substring(Fname.IndexOf('T'));



                        //string originalFileName = subStringTEST_.Substring(subStringTEST_.IndexOf('P'));




                        textBox5.Text = fName;

                        string erid = getStringBetweenString(fName, "M.", "_");



                        textBox1.Text = erid;
                        ERID = erid;

                        //get group name (without carriers) from erid

                        GetGroupName getGroupName = new GetGroupName();
                        getGroupName.GetGroupNameMethod(ERID);
                        string groupName = GetGroupName.GroupName;

                        this.Text = "FC Helper - " + groupName;
                        //get client ID
                        //GetClientID getClientID = new GetClientID();
                        //getClientID.GetClientIDMethod(groupName);


                        //ClientID = GetClientID.ClientID;

                        //textBox2.Text = ClientID;









                        //get group folder file path
                        string groupFolderFilePath = @"C:\Users\14025\Documents\File Consultants\Groups\" + groupName;
                        string[] groupFolderFilePathOfFile = Directory.GetFiles(groupFolderFilePath, "*" + fName + ".*", SearchOption.AllDirectories);

                        textBox6.Text = "NA";
                        listView1.Items.Add("File has not processed yet.");



                        foreach (string file in groupFolderFilePathOfFile)
                        {

                            if (!File.Exists(file))
                            {
                                textBox6.Text = "";

                            }
                            else if (File.Exists(file))
                            {
                                listView1.Clear();
                                textBox6.Clear();
                                NoDateTimeStampsGiven = true;

                                string fileName = Path.GetFileName(file);

                                Fname = null;
                                Fname = fileName;

                                listView1.Items.Add(fileName);



                                textBox3.Text = fileName;
                                string dateStamp = getStringBetweenString(fileName, "", "_");
                                string timeStamp = getStringBetweenString(fileName, "_", "_");

                                string processedDateTimeStamp = dateStamp + "_" + timeStamp;


                                textBox6.Text = processedDateTimeStamp;

                                string destinationPath = @"C:\Users\14025\Documents\File Consultants\Brandon\Temp\" + fileName;
                                //read file in RichTextbox and Copy/Paste the file in question
                                if (!File.Exists(destinationPath))
                                {
                                    File.Copy(file, destinationPath);
                                    listView1.Items.Add(textBox5.Text);
                                    string fileText = File.ReadAllText(destinationPath);
                                    richTextBox1.Text = fileText;
                                }
                                else
                                {
                                    MessageBox.Show("This file already exists in the Temp folder.", "Can't Do It. Wouldn't be Proper.", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                                }



                                //search for the groups error folder error reports
                                GetErrorReportFileName name = new GetErrorReportFileName();
                                name.GetErrorReportFileNameMethod(erid, processedDateTimeStamp);

                                //search for the groups error folder for psvs
                                GetInputPSVFileNames names = new GetInputPSVFileNames();
                                names.GetInputPSVFileNamesMethod(erid, processedDateTimeStamp);

                                //copy error reports to temp folder



                            }



                        }
                        foreach (string groupFolderItem in GetErrorReportFileName.ErrorReportGroupFolderLocations)
                        {

                            //get source path+filenames
                            string sourceFile = groupFolderItem;


                            string errorReportfilename = Path.GetFileName(sourceFile);
                            string destinationFile = @"C:\Users\14025\Documents\File Consultants\Brandon\Temp\" + errorReportfilename;

                            if (!File.Exists(destinationFile))
                            {
                                File.Copy(sourceFile, destinationFile);

                            }
                        }

                        //copy psv's to temp folder
                        foreach (string groupFolderItem in GetInputPSVFileNames.PSVFolderLocations)
                        {


                            //get source path+filenames
                            string sourceFile = groupFolderItem;


                            string psvfilename = Path.GetFileName(sourceFile);
                            string destinationFile = @"C:\Users\14025\Documents\File Consultants\Brandon\Temp\" + psvfilename;

                            if (!File.Exists(destinationFile))
                            {
                                File.Copy(sourceFile, destinationFile);

                            }



                        }

                        //add error report to listview2

                        foreach (string file1 in GetErrorReportFileName.ErrorReportFileNamesOnly)
                        {
                            listView2.Items.Add(file1);

                        }
                        //add inputpsvs to listview3

                        foreach (string file1 in GetInputPSVFileNames.InputPSVFileNamesOnly)
                        {

                            listView3.Items.Add(file1);

                        }
                        //add outputpsvs to listview4

                        foreach (string file1 in GetInputPSVFileNames.OutputPSVFileNamesOnly)
                        {
                            //MessageBox.Show(file);

                            listView4.Items.Add(file1);

                        }

                    }
                    else if (fName.Contains("PF.PFM") && !fName.Contains("PF_TEST"))//standard PF.PFM. production file filename or the mistaken PF.PFM_
                    {


                        string origFileNameWithPFPFM = fName.Substring(fName.IndexOf('P'));



                        textBox5.Text = origFileNameWithPFPFM;

                        string erid = getStringBetweenString(origFileNameWithPFPFM, "M.", "_");



                        string dateStamp = getStringBetweenString(fName, "", "_");
                        string timeStamp = getStringBetweenString(fName, "_", "_");


                        string processedDateTimeStamp = dateStamp + "_" + timeStamp;

                        textBox1.Text = erid;
                        ERID = erid;

                        //get group name (without carriers) from erid

                        GetGroupName getGroupName = new GetGroupName();
                        getGroupName.GetGroupNameMethod(ERID);
                        string groupName = GetGroupName.GroupName;

                        this.Text = "FC Helper - " + groupName;

                        //get client ID
                        //GetClientID getClientID = new GetClientID();
                        //getClientID.GetClientIDMethod(groupName);


                        //ClientID = GetClientID.ClientID;

                        //textBox2.Text = ClientID;
                        textBox6.Text = processedDateTimeStamp;


                        //get group folder file path
                        string groupFolderFilePath = @"C:\Users\14025\Documents\File Consultants\Groups\" + groupName;
                        string[] groupFolderFilePathOfFile = Directory.GetFiles(groupFolderFilePath, Fname, SearchOption.AllDirectories);

                        foreach (string file in groupFolderFilePathOfFile)
                        {

                            //find our file only


                            //string filePath = System.IO.Path.GetDirectoryName(file);
                            //string fileName = System.IO.Path.GetFileName(file);
                            string destinationPath = @"C:\Users\14025\Documents\File Consultants\Brandon\Temp\" + fName;
                            //read file in RichTextbox and Copy/Paste the file in question
                            if (!File.Exists(destinationPath))
                            {
                                File.Copy(file, destinationPath);
                                string fileText = File.ReadAllText(destinationPath);
                                richTextBox1.Text = fileText;


                            }
                            else
                            {
                                MessageBox.Show("This file already exists in the Temp folder.", "Can't Do It. Wouldn't be Proper.", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);

                                string fileText = File.ReadAllText(destinationPath);
                                richTextBox1.Text = fileText;
                            }




                        }

                        listView1.Items.Add(textBox5.Text);




                        //search for the groups error folder error reports
                        GetErrorReportFileName name = new GetErrorReportFileName();
                        name.GetErrorReportFileNameMethod(ERID, processedDateTimeStamp);

                        //search for the groups error folder for psvs
                        GetInputPSVFileNames names = new GetInputPSVFileNames();
                        names.GetInputPSVFileNamesMethod(ERID, processedDateTimeStamp);

                        //copy error reports to temp folder

                        foreach (string groupFolderItem in GetErrorReportFileName.ErrorReportGroupFolderLocations)
                        {

                            //get source path+filenames
                            string sourceFile = groupFolderItem;


                            string errorReportfilename = Path.GetFileName(sourceFile);
                            string destinationFile = @"C:\Users\14025\Documents\File Consultants\Brandon\Temp\" + errorReportfilename;

                            if (!File.Exists(destinationFile))
                            {
                                File.Copy(sourceFile, destinationFile);

                            }

                        }

                        //copy psv's to temp folder
                        foreach (string groupFolderItem in GetInputPSVFileNames.PSVFolderLocations)
                        {


                            //get source path+filenames
                            string sourceFile = groupFolderItem;


                            string psvfilename = Path.GetFileName(sourceFile);
                            string destinationFile = @"C:\Users\14025\Documents\File Consultants\Brandon\Temp\" + psvfilename;

                            if (!File.Exists(destinationFile))
                            {
                                File.Copy(sourceFile, destinationFile);

                            }
                        }

                        //add error report to listview2

                        foreach (string file in GetErrorReportFileName.ErrorReportFileNamesOnly)
                        {
                            listView2.Items.Add(file);

                        }
                        //add inputpsvs to listview3

                        foreach (string file in GetInputPSVFileNames.InputPSVFileNamesOnly)
                        {

                            listView3.Items.Add(file);

                        }
                        //add outputpsvs to listview4

                        foreach (string file in GetInputPSVFileNames.OutputPSVFileNamesOnly)
                        {
                            //MessageBox.Show(file);

                            listView4.Items.Add(file);

                        }



                    }
                    else if (fName.Contains("PF_TEST") && fName.Contains("PF.PFM"))//standard test file
                    {

                        //Fname = fName;

                        string origFileNameWithPF_TEST = fName.Substring(fName.IndexOf('P'));


                        string subStringTEST_ = origFileNameWithPF_TEST.Substring(origFileNameWithPF_TEST.IndexOf('T'));



                        string originalFileName = subStringTEST_.Substring(subStringTEST_.IndexOf('P'));




                        textBox5.Text = originalFileName;

                        string erid = getStringBetweenString(originalFileName, "M.", "_");
                        string processedDateTimeStampWithExcess1 = getStringBetweenString(fName, "", "_2");
                        string processedDateTimeStampWithExcess2 = getStringBetweenString(fName, "", "_P");


                        string dateStamp = getStringBetweenString(processedDateTimeStampWithExcess2, "", "_");
                        string timeStamp = getStringBetweenString(processedDateTimeStampWithExcess2, "_", "_");

                        string processedDateTimeStamp = dateStamp + "_" + timeStamp;


                        textBox1.Text = erid;
                        ERID = erid;

                        //get group name (without carriers) from erid

                        GetGroupName getGroupName = new GetGroupName();
                        getGroupName.GetGroupNameMethod(ERID);
                        string groupName = GetGroupName.GroupName;


                        this.Text = "FC Helper - " + groupName;
                        //get client ID
                        //GetClientID getClientID = new GetClientID();
                        //getClientID.GetClientIDMethod(groupName);


                        //ClientID = GetClientID.ClientID;

                        //textBox2.Text = ClientID;
                        textBox6.Text = processedDateTimeStamp;


                        //get group folder file path
                        string groupFolderFilePath = @"C:\Users\14025\Documents\File Consultants\Groups\" + groupName;
                        string[] groupFolderFilePathOfFile = Directory.GetFiles(groupFolderFilePath, Fname, SearchOption.AllDirectories);

                        foreach (string file in groupFolderFilePathOfFile)
                        {
                            //read file in RichTextbox
                            string filePath = System.IO.Path.GetDirectoryName(Fname);
                            string fileName = System.IO.Path.GetFileName(Fname);
                            string destinationPath = @"C:\Users\14025\Documents\File Consultants\Brandon\Temp\" + fName;
                            //read file in RichTextbox and Copy/Paste the file in question
                            if (!File.Exists(destinationPath))
                            {
                                File.Copy(file, destinationPath);
                                listView1.Items.Add(textBox5.Text);
                                string fileText = File.ReadAllText(destinationPath);
                                richTextBox1.Text = fileText;
                            }
                            else
                            {
                                MessageBox.Show("This file already exists in the Temp folder.", "Can't Do It. Wouldn't be Proper.", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                            }

                        }



                        //search for the groups error folder error reports
                        GetErrorReportFileName name = new GetErrorReportFileName();
                        name.GetErrorReportFileNameMethod(erid, processedDateTimeStamp);

                        //search for the groups error folder for psvs
                        GetInputPSVFileNames names = new GetInputPSVFileNames();
                        names.GetInputPSVFileNamesMethod(erid, processedDateTimeStamp);

                        //copy error reports to temp folder

                        foreach (string groupFolderItem in GetErrorReportFileName.ErrorReportGroupFolderLocations)
                        {

                            //get source path+filenames
                            string sourceFile = groupFolderItem;


                            string errorReportfilename = Path.GetFileName(sourceFile);
                            string destinationFile = @"C:\Users\14025\Documents\File Consultants\Brandon\Temp\" + errorReportfilename;

                            if (!File.Exists(destinationFile))
                            {
                                File.Copy(sourceFile, destinationFile);

                            }
                        }

                        //copy psv's to temp folder
                        foreach (string groupFolderItem in GetInputPSVFileNames.PSVFolderLocations)
                        {


                            //get source path+filenames
                            string sourceFile = groupFolderItem;


                            string psvfilename = Path.GetFileName(sourceFile);
                            string destinationFile = @"C:\Users\14025\Documents\File Consultants\Brandon\Temp\" + psvfilename;

                            if (!File.Exists(destinationFile))
                            {
                                File.Copy(sourceFile, destinationFile);

                            }



                        }

                        //add error report to listview2

                        foreach (string file in GetErrorReportFileName.ErrorReportFileNamesOnly)
                        {
                            listView2.Items.Add(file);

                        }
                        //add inputpsvs to listview3

                        foreach (string file in GetInputPSVFileNames.InputPSVFileNamesOnly)
                        {

                            listView3.Items.Add(file);

                        }
                        //add outputpsvs to listview4

                        foreach (string file in GetInputPSVFileNames.OutputPSVFileNamesOnly)
                        {
                            //MessageBox.Show(file);

                            listView4.Items.Add(file);

                        }
                    }
                    else if (fName.Contains(".xlsm"))//standard error report file
                    {


                        //get date/time stamp

                        string dateStamp = getStringBetweenString(fName, "", "_");
                        string timeStamp = getStringBetweenString(fName, "_", "_");

                        string processedDateTimeStamp = dateStamp + "_" + timeStamp;


                        //get employer name

                        string fNameMinusDate = fName.Substring(fName.IndexOf('_'));
                        string fNameMinusTime = getStringBetweenString(fNameMinusDate, "_", ".xlsm");

                        string employerName = getStringBetweenString(fNameMinusTime, "_", "_");


                        //search groups folder for employer name

                        try
                        {
                            string[] employerFolderNames = Directory.GetDirectories(GroupsFolderPath, employerName + "*", SearchOption.TopDirectoryOnly);

                            foreach (string employerFolder in employerFolderNames)
                            {
                                string groupName1 = Path.GetFileName(employerFolder);
                                string strToCompare = "-";

                                if (groupName1.Split().Count(r => r == strToCompare) > 1)
                                {
                                    //do nothing and go to the next iteration
                                }
                                else
                                {
                                    GroupName = groupName1;//delivers goupName without carriers
                                                           //MessageBox.Show(GroupName);
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show(ex.ToString());

                        }


                        //search for date/time stamp

                        string[] filesInEmployerFolder = Directory.GetFiles(GroupsFolderPath + GroupName + @"\Import\Done", processedDateTimeStamp + "*", SearchOption.AllDirectories);
                        foreach (string filePath in filesInEmployerFolder)
                        {
                            //MessageBox.Show(filePath);
                            string file = Path.GetFileName(filePath);
                            textBox3.Text = file;
                            Fname = file;


                            string destinationPath = @"C:\Users\14025\Documents\File Consultants\Brandon\Temp\" + file;
                            //read file in RichTextbox and Copy/Paste the file in question
                            if (!File.Exists(destinationPath))
                            {
                                File.Copy(filePath, destinationPath);
                                string fileText = File.ReadAllText(destinationPath);
                                richTextBox1.Text = fileText;


                            }
                            else
                            {
                                MessageBox.Show("This file already exists in the Temp folder.", "Can't Do It. Wouldn't be Proper.", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);

                                string fileText = File.ReadAllText(destinationPath);
                                richTextBox1.Text = fileText;
                            }
                        }
                        string origFileNameWithPFPFM = Fname.Substring(Fname.IndexOf('P'));



                        textBox5.Text = origFileNameWithPFPFM;

                        string erid = getStringBetweenString(origFileNameWithPFPFM, "M.", "_");





                        textBox1.Text = erid;
                        ERID = erid;

                        //get group name (without carriers) from erid

                        GetGroupName getGroupName = new GetGroupName();
                        getGroupName.GetGroupNameMethod(ERID);
                        string groupName = GetGroupName.GroupName;

                        this.Text = "FC Helper - " + groupName;

                        //get client ID
                        //GetClientID getClientID = new GetClientID();
                        //getClientID.GetClientIDMethod(groupName);


                        //ClientID = GetClientID.ClientID;

                        //textBox2.Text = ClientID;
                        textBox6.Text = processedDateTimeStamp;


                        //get group folder file path
                        //string groupFolderFilePath = @"C:\Users\14025\Documents\File Consultants\Groups\" + groupName;
                        //string[] groupFolderFilePathOfFile = Directory.GetFiles(groupFolderFilePath, Fname, SearchOption.AllDirectories);



                        listView1.Items.Add(textBox5.Text);




                        //search for the groups error folder error reports
                        GetErrorReportFileName name = new GetErrorReportFileName();
                        name.GetErrorReportFileNameMethod(ERID, processedDateTimeStamp);

                        //search for the groups error folder for psvs
                        GetInputPSVFileNames names = new GetInputPSVFileNames();
                        names.GetInputPSVFileNamesMethod(ERID, processedDateTimeStamp);

                        //copy error reports to temp folder

                        foreach (string groupFolderItem in GetErrorReportFileName.ErrorReportGroupFolderLocations)
                        {

                            //get source path+filenames
                            string sourceFile = groupFolderItem;


                            string errorReportfilename = Path.GetFileName(sourceFile);
                            string destinationFile = @"C:\Users\14025\Documents\File Consultants\Brandon\Temp\" + errorReportfilename;

                            if (!File.Exists(destinationFile))
                            {
                                File.Copy(sourceFile, destinationFile);

                            }

                        }

                        //copy psv's to temp folder
                        foreach (string groupFolderItem in GetInputPSVFileNames.PSVFolderLocations)
                        {


                            //get source path+filenames
                            string sourceFile = groupFolderItem;


                            string psvfilename = Path.GetFileName(sourceFile);
                            string destinationFile = @"C:\Users\14025\Documents\File Consultants\Brandon\Temp\" + psvfilename;

                            if (!File.Exists(destinationFile))
                            {
                                File.Copy(sourceFile, destinationFile);

                            }
                        }

                        //add error report to listview2

                        foreach (string file in GetErrorReportFileName.ErrorReportFileNamesOnly)
                        {
                            listView2.Items.Add(file);

                        }
                        //add inputpsvs to listview3

                        foreach (string file in GetInputPSVFileNames.InputPSVFileNamesOnly)
                        {

                            listView3.Items.Add(file);

                        }
                        //add outputpsvs to listview4

                        foreach (string file in GetInputPSVFileNames.OutputPSVFileNamesOnly)
                        {
                            //MessageBox.Show(file);

                            listView4.Items.Add(file);

                        }
                    }
                    
                }


            
            }
            catch (System.Exception ex)
            {

                MessageBox.Show("Method: Form_Main_DragDrop\r Something prevented your file from opening.\r\r" + ex);
            }
            textBox1.Focus();
            textBox1.SelectAll();
            textBox1.Copy();
            Fname = textBox3.Text;  
        }//drag drop file

        private void textBox1_KeyPress(object sender, KeyPressEventArgs e)//textbox 1 enter
        {
            if (textBox1.Text != "" && e.KeyChar == (char)Keys.Return)
            {
                button2.PerformClick();
            }

        }

        private void textBox3_KeyPress(object sender, KeyPressEventArgs e)
        {

            //GetGroupName.GroupName = null;

            if (TextBox1Enter == true && e.KeyChar == (char)Keys.Return)
            {
                //Fname = null;
                string fName = textBox3.Text;

                //get erid from filename



                int countUnderscores = fName.Count(x => x == '_');//counting the number of underscores in filename
                int countPeriods = fName.Count(x => x == '.');//counting the number of periods in filename

                int totalFilenameCharacterLength = fName.Length;//

                int firstLetterIndex = fName.IndexOf('P');

                Fname = fName;
                textBox3.Text = fName;



                if (firstLetterIndex == 0 && fName.Contains("PF_TEST"))
                {


                    //MessageBox.Show(@"firstLetterIndex == 0 && fName.Contains(PF_TEST)");

                    string subStringTEST_ = fName.Substring(fName.IndexOf('T'));



                    string originalFileName = subStringTEST_.Substring(subStringTEST_.IndexOf('P'));





                    string erid = getStringBetweenString(originalFileName, "M.", "_");



                    textBox1.Text = erid;
                    ERID = erid;

                    //get group name (without carriers) from erid

                    GetGroupName getGroupName = new GetGroupName();
                    getGroupName.GetGroupNameMethod(ERID);
                    string groupName = GetGroupName.GroupName;

                    this.Text = "FC Helper - " + groupName;
                    //get client ID
                    //GetClientID getClientID = new GetClientID();
                    //getClientID.GetClientIDMethod(groupName);


                    //ClientID = GetClientID.ClientID;

                    //textBox2.Text = ClientID;









                    //get group folder file path
                    string groupFolderFilePath = @"C:\Users\14025\Documents\File Consultants\Groups\" + groupName;

                    try
                    {
                        string[] groupFolderFilePathOfFile = Directory.GetFiles(groupFolderFilePath, "*" + Fname + ".*", SearchOption.AllDirectories);

                        textBox6.Text = "NA";
                        textBox5.Text = originalFileName;
                        listView1.Items.Add("File has not processed yet.");

                        foreach (string file in groupFolderFilePathOfFile)
                        {


                            if (!File.Exists(file))
                            {
                                textBox6.Text = "NA";
                                textBox5.Text = originalFileName;
                                listView1.Items.Add("File has not processed yet.");
                                listView2.Items.Add("File has not processed yet.");
                                listView3.Items.Add("File has not processed yet.");
                            }
                            else if (File.Exists(file))
                            {

                                listView1.Clear();
                                textBox6.Clear();

                                string fileName = Path.GetFileName(file);
                                textBox3.Text = fileName;
                                //textBox5.Text = originalFileName;
                                Fname = fileName;
                                string processedDateTimeStampWithExcess1 = getStringBetweenString(fileName, "", "_2");
                                string processedDateTimeStampWithExcess2 = getStringBetweenString(fileName, "", "_P");


                                string dateStamp = getStringBetweenString(processedDateTimeStampWithExcess2, "", "_");
                                string timeStamp = getStringBetweenString(processedDateTimeStampWithExcess2, "_", "_");

                                string processedDateTimeStamp = dateStamp + "_" + timeStamp;


                                textBox6.Text = processedDateTimeStamp;

                                string destinationPath = @"C:\Users\14025\Documents\File Consultants\Brandon\Temp\" + fName;
                                //read file in RichTextbox and Copy/Paste the file in question
                                if (!File.Exists(destinationPath))
                                {
                                    File.Copy(file, destinationPath);
                                    listView1.Items.Add(textBox5.Text);
                                    string fileText = File.ReadAllText(destinationPath);
                                    richTextBox1.Text = fileText;
                                }
                                else
                                {
                                    MessageBox.Show("This file already exists in the Temp folder.", "Can't Do It. Wouldn't be Proper.", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                                }



                                //search for the groups error folder error reports
                                GetErrorReportFileName name = new GetErrorReportFileName();
                                name.GetErrorReportFileNameMethod(erid, processedDateTimeStamp);

                                //search for the groups error folder for psvs
                                GetInputPSVFileNames names = new GetInputPSVFileNames();
                                names.GetInputPSVFileNamesMethod(erid, processedDateTimeStamp);

                                //copy error reports to temp folder

                                foreach (string groupFolderItem in GetErrorReportFileName.ErrorReportGroupFolderLocations)
                                {

                                    //get source path+filenames
                                    string sourceFile = groupFolderItem;


                                    string errorReportfilename = Path.GetFileName(sourceFile);
                                    string destinationFile = @"C:\Users\14025\Documents\File Consultants\Brandon\Temp\" + errorReportfilename;

                                    if (!File.Exists(destinationFile))
                                    {
                                        File.Copy(sourceFile, destinationFile);

                                    }
                                }

                                //copy psv's to temp folder
                                foreach (string groupFolderItem in GetInputPSVFileNames.PSVFolderLocations)
                                {


                                    //get source path+filenames
                                    string sourceFile = groupFolderItem;


                                    string psvfilename = Path.GetFileName(sourceFile);
                                    string destinationFile = @"C:\Users\14025\Documents\File Consultants\Brandon\Temp\" + psvfilename;

                                    if (!File.Exists(destinationFile))
                                    {
                                        File.Copy(sourceFile, destinationFile);

                                    }



                                }

                                //add error report to listview2

                                foreach (string file1 in GetErrorReportFileName.ErrorReportFileNamesOnly)
                                {
                                    listView2.Items.Add(file1);

                                }
                                //add inputpsvs to listview3

                                foreach (string file1 in GetInputPSVFileNames.InputPSVFileNamesOnly)
                                {

                                    listView3.Items.Add(file1);

                                }
                                //add outputpsvs to listview4

                                foreach (string file1 in GetInputPSVFileNames.OutputPSVFileNamesOnly)
                                {
                                    //MessageBox.Show(file);

                                    listView4.Items.Add(file1);

                                }

                            }
                        }
                    }
                    catch (ArgumentNullException ex)
                    {
                        MessageBox.Show("No file found.");
                    }


                   



                    


                }
                else if (firstLetterIndex == 0)
                {

                    //Fname = fName;

                    //string origFileNameWithPF_TEST = fName.Substring(fName.IndexOf('P'));


                    //string subStringTEST_ = Fname.Substring(Fname.IndexOf('T'));



                    //string originalFileName = subStringTEST_.Substring(subStringTEST_.IndexOf('P'));




                    textBox5.Text = fName;

                    string erid = getStringBetweenString(Fname, "M.", "_");



                    textBox1.Text = erid;
                    ERID = erid;

                    //get group name (without carriers) from erid

                    GetGroupName getGroupName = new GetGroupName();
                    getGroupName.GetGroupNameMethod(ERID);
                    string groupName = GetGroupName.GroupName;

                    this.Text = "FC Helper - " + groupName;
                    //get client ID
                    //GetClientID getClientID = new GetClientID();
                    //getClientID.GetClientIDMethod(groupName);


                    //ClientID = GetClientID.ClientID;

                    //textBox2.Text = ClientID;









                    //get group folder file path
                    string groupFolderFilePath = @"C:\Users\14025\Documents\File Consultants\Groups\" + groupName;
                    string[] groupFolderFilePathOfFile = Directory.GetFiles(groupFolderFilePath, "*" + Fname + ".*", SearchOption.AllDirectories);

                    textBox6.Text = "NA";
                    listView1.Items.Add("File has not processed yet.");



                    foreach (string file in groupFolderFilePathOfFile)
                    {

                        if (!File.Exists(file))
                        {
                            textBox6.Text = "";

                        }
                        else if (File.Exists(file))
                        {
                            listView1.Clear();
                            textBox6.Clear();
                            NoDateTimeStampsGiven = true;

                            listView1.Items.Add(Fname);

                            string fileName = Path.GetFileName(file);
                            Fname = fileName;
                            textBox3.Text = fileName;
                            string dateStamp = getStringBetweenString(fileName, "", "_");
                            string timeStamp = getStringBetweenString(fileName, "_", "_");

                            string processedDateTimeStamp = dateStamp + "_" + timeStamp;


                            textBox6.Text = processedDateTimeStamp;

                            string destinationPath = @"C:\Users\14025\Documents\File Consultants\Brandon\Temp\" + fileName;
                            //read file in RichTextbox and Copy/Paste the file in question
                            if (!File.Exists(destinationPath))
                            {
                                File.Copy(file, destinationPath);
                                listView1.Items.Add(textBox5.Text);
                                string fileText = File.ReadAllText(destinationPath);
                                richTextBox1.Text = fileText;
                            }
                            else
                            {
                                MessageBox.Show("This file already exists in the Temp folder.", "Can't Do It. Wouldn't be Proper.", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                            }



                            //search for the groups error folder error reports
                            GetErrorReportFileName name = new GetErrorReportFileName();
                            name.GetErrorReportFileNameMethod(erid, processedDateTimeStamp);

                            //search for the groups error folder for psvs
                            GetInputPSVFileNames names = new GetInputPSVFileNames();
                            names.GetInputPSVFileNamesMethod(erid, processedDateTimeStamp);

                            //copy error reports to temp folder



                        }



                    }
                    foreach (string groupFolderItem in GetErrorReportFileName.ErrorReportGroupFolderLocations)
                    {

                        //get source path+filenames
                        string sourceFile = groupFolderItem;


                        string errorReportfilename = Path.GetFileName(sourceFile);
                        string destinationFile = @"C:\Users\14025\Documents\File Consultants\Brandon\Temp\" + errorReportfilename;

                        if (!File.Exists(destinationFile))
                        {
                            File.Copy(sourceFile, destinationFile);

                        }
                    }

                    //copy psv's to temp folder
                    foreach (string groupFolderItem in GetInputPSVFileNames.PSVFolderLocations)
                    {


                        //get source path+filenames
                        string sourceFile = groupFolderItem;


                        string psvfilename = Path.GetFileName(sourceFile);
                        string destinationFile = @"C:\Users\14025\Documents\File Consultants\Brandon\Temp\" + psvfilename;

                        if (!File.Exists(destinationFile))
                        {
                            File.Copy(sourceFile, destinationFile);

                        }



                    }

                    //add error report to listview2

                    foreach (string file1 in GetErrorReportFileName.ErrorReportFileNamesOnly)
                    {
                        listView2.Items.Add(file1);

                    }
                    //add inputpsvs to listview3

                    foreach (string file1 in GetInputPSVFileNames.InputPSVFileNamesOnly)
                    {

                        listView3.Items.Add(file1);

                    }
                    //add outputpsvs to listview4

                    foreach (string file1 in GetInputPSVFileNames.OutputPSVFileNamesOnly)
                    {
                        //MessageBox.Show(file);

                        listView4.Items.Add(file1);

                    }

                }
                else if (fName.Contains("PF.PFM") && !fName.Contains("PF_TEST"))//standard PF.PFM. production file filename or the mistaken PF.PFM_
                {


                    string origFileNameWithPFPFM = fName.Substring(fName.IndexOf('P'));



                    textBox5.Text = origFileNameWithPFPFM;

                    string erid = getStringBetweenString(origFileNameWithPFPFM, "M.", "_");



                    string dateStamp = getStringBetweenString(fName, "", "_");
                    string timeStamp = getStringBetweenString(fName, "_", "_");


                    string processedDateTimeStamp = dateStamp + "_" + timeStamp;

                    textBox1.Text = erid;
                    ERID = erid;

                    //get group name (without carriers) from erid

                    GetGroupName getGroupName = new GetGroupName();
                    getGroupName.GetGroupNameMethod(ERID);
                    string groupName = GetGroupName.GroupName;


                    this.Text = "FC Helper - " + groupName;
                    //get client ID
                    //GetClientID getClientID = new GetClientID();
                    //getClientID.GetClientIDMethod(groupName);


                    //ClientID = GetClientID.ClientID;

                    //textBox2.Text = ClientID;
                    textBox6.Text = processedDateTimeStamp;


                    //get group folder file path
                    string groupFolderFilePath = @"C:\Users\14025\Documents\File Consultants\Groups\" + groupName;
                    string[] groupFolderFilePathOfFile = Directory.GetFiles(groupFolderFilePath, fName, SearchOption.AllDirectories);

                    foreach (string file in groupFolderFilePathOfFile)
                    {
                       // MessageBox.Show("file path: "+file);
                        string fileName = Path.GetFileName(file);
                        Fname = fileName;
                        //MessageBox.Show("file name: " + fileName);

                        //find our file only


                        //string filePath = System.IO.Path.GetDirectoryName(file);
                        //string fileName = System.IO.Path.GetFileName(file);
                        string destinationPath = @"C:\Users\14025\Documents\File Consultants\Brandon\Temp\" + fName;
                        //read file in RichTextbox and Copy/Paste the file in question
                        if (!File.Exists(destinationPath))
                        {
                            //File.Copy(file, destinationPath);
                            //string fileText = File.ReadAllText(destinationPath);
                            //richTextBox1.Text = fileText;
                            textBox3.Text = fName;
                            textBox5.Text = fName;
                            string fileText = File.ReadAllText(destinationPath);
                            richTextBox1.Text = fileText;
                            textBox6.Text = "N/A";
                            listView1.Items.Add("This file is not in this employer's folder.");


                        }
                        else
                        {
                            //MessageBox.Show("This file already exists in the Temp folder.", "Can't Do It. Wouldn't be Proper.", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                            
                            string fileText = File.ReadAllText(destinationPath);
                            richTextBox1.Text = fileText;
                        }




                    }
                    //Fname = null;
                    Fname = textBox3.Text;
                    listView1.Items.Add(textBox5.Text);




                    //search for the groups error folder error reports
                    GetErrorReportFileName name = new GetErrorReportFileName();
                    name.GetErrorReportFileNameMethod(ERID, processedDateTimeStamp);

                    //search for the groups error folder for psvs
                    GetInputPSVFileNames names = new GetInputPSVFileNames();
                    names.GetInputPSVFileNamesMethod(ERID, processedDateTimeStamp);

                    //copy error reports to temp folder

                    foreach (string groupFolderItem in GetErrorReportFileName.ErrorReportGroupFolderLocations)
                    {

                        //get source path+filenames
                        string sourceFile = groupFolderItem;


                        string errorReportfilename = Path.GetFileName(sourceFile);
                        string destinationFile = @"C:\Users\14025\Documents\File Consultants\Brandon\Temp\" + errorReportfilename;

                        if (!File.Exists(destinationFile))
                        {
                            File.Copy(sourceFile, destinationFile);

                        }

                    }

                    //copy psv's to temp folder
                    foreach (string groupFolderItem in GetInputPSVFileNames.PSVFolderLocations)
                    {


                        //get source path+filenames
                        string sourceFile = groupFolderItem;


                        string psvfilename = Path.GetFileName(sourceFile);
                        string destinationFile = @"C:\Users\14025\Documents\File Consultants\Brandon\Temp\" + psvfilename;

                        if (!File.Exists(destinationFile))
                        {
                            File.Copy(sourceFile, destinationFile);

                        }
                    }

                    //add error report to listview2

                    foreach (string file in GetErrorReportFileName.ErrorReportFileNamesOnly)
                    {
                        listView2.Items.Add(file);

                    }
                    //add inputpsvs to listview3

                    foreach (string file in GetInputPSVFileNames.InputPSVFileNamesOnly)
                    {

                        listView3.Items.Add(file);

                    }
                    //add outputpsvs to listview4

                    foreach (string file in GetInputPSVFileNames.OutputPSVFileNamesOnly)
                    {
                        //MessageBox.Show(file);

                        listView4.Items.Add(file);

                    }



                }
                else if (fName.Contains("PF_TEST") && fName.Contains("PF.PFM"))//standard test file
                {

                    //Fname = fName;

                    string origFileNameWithPF_TEST = fName.Substring(fName.IndexOf('P'));


                    string subStringTEST_ = origFileNameWithPF_TEST.Substring(origFileNameWithPF_TEST.IndexOf('T'));



                    string originalFileName = subStringTEST_.Substring(subStringTEST_.IndexOf('P'));




                    textBox5.Text = originalFileName;

                    string erid = getStringBetweenString(originalFileName, "M.", "_");
                    string processedDateTimeStampWithExcess1 = getStringBetweenString(fName, "", "_2");
                    string processedDateTimeStampWithExcess2 = getStringBetweenString(fName, "", "_P");


                    string dateStamp = getStringBetweenString(processedDateTimeStampWithExcess2, "", "_");
                    string timeStamp = getStringBetweenString(processedDateTimeStampWithExcess2, "_", "_");

                    string processedDateTimeStamp = dateStamp + "_" + timeStamp;


                    textBox1.Text = erid;
                    ERID = erid;

                    //get group name (without carriers) from erid

                    GetGroupName getGroupName = new GetGroupName();
                    getGroupName.GetGroupNameMethod(ERID);
                    string groupName = GetGroupName.GroupName;


                    this.Text = "FC Helper - " + groupName;
                    //get client ID
                    //GetClientID getClientID = new GetClientID();
                    //getClientID.GetClientIDMethod(groupName);


                    //ClientID = GetClientID.ClientID;

                    //textBox2.Text = ClientID;
                    textBox6.Text = processedDateTimeStamp;


                    //get group folder file path
                    string groupFolderFilePath = @"C:\Users\14025\Documents\File Consultants\Groups\" + groupName;
                    string[] groupFolderFilePathOfFile = Directory.GetFiles(groupFolderFilePath, fName, SearchOption.AllDirectories);

                    foreach (string file in groupFolderFilePathOfFile)
                    {
                        //read file in RichTextbox
                        string filePath = System.IO.Path.GetDirectoryName(fName);
                        string fileName = System.IO.Path.GetFileName(fName);
                        Fname = fileName;
                        string destinationPath = @"C:\Users\14025\Documents\File Consultants\Brandon\Temp\" + fileName;
                        //read file in RichTextbox and Copy/Paste the file in question
                        if (!File.Exists(destinationPath))
                        {
                            File.Copy(file, destinationPath);
                            listView1.Items.Add(textBox5.Text);
                            string fileText = File.ReadAllText(destinationPath);
                            richTextBox1.Text = fileText;
                        }
                        else
                        {
                            MessageBox.Show("This file already exists in the Temp folder.", "Can't Do It. Wouldn't be Proper.", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                        }

                    }

                    Fname = null;
                    Fname = textBox3.Text;

                    //search for the groups error folder error reports
                    GetErrorReportFileName name = new GetErrorReportFileName();
                    name.GetErrorReportFileNameMethod(erid, processedDateTimeStamp);

                    //search for the groups error folder for psvs
                    GetInputPSVFileNames names = new GetInputPSVFileNames();
                    names.GetInputPSVFileNamesMethod(erid, processedDateTimeStamp);

                    //copy error reports to temp folder

                    foreach (string groupFolderItem in GetErrorReportFileName.ErrorReportGroupFolderLocations)
                    {

                        //get source path+filenames
                        string sourceFile = groupFolderItem;


                        string errorReportfilename = Path.GetFileName(sourceFile);
                        string destinationFile = @"C:\Users\14025\Documents\File Consultants\Brandon\Temp\" + errorReportfilename;

                        if (!File.Exists(destinationFile))
                        {
                            File.Copy(sourceFile, destinationFile);

                        }
                    }

                    //copy psv's to temp folder
                    foreach (string groupFolderItem in GetInputPSVFileNames.PSVFolderLocations)
                    {


                        //get source path+filenames
                        string sourceFile = groupFolderItem;


                        string psvfilename = Path.GetFileName(sourceFile);
                        string destinationFile = @"C:\Users\14025\Documents\File Consultants\Brandon\Temp\" + psvfilename;

                        if (!File.Exists(destinationFile))
                        {
                            File.Copy(sourceFile, destinationFile);

                        }



                    }

                    //add error report to listview2

                    foreach (string file in GetErrorReportFileName.ErrorReportFileNamesOnly)
                    {
                        listView2.Items.Add(file);

                    }
                    //add inputpsvs to listview3

                    foreach (string file in GetInputPSVFileNames.InputPSVFileNamesOnly)
                    {

                        listView3.Items.Add(file);

                    }
                    //add outputpsvs to listview4

                    foreach (string file in GetInputPSVFileNames.OutputPSVFileNamesOnly)
                    {
                        //MessageBox.Show(file);

                        listView4.Items.Add(file);

                    }
                }
                else if (fName.Contains(".xlsm"))//standard error report file
                {


                    //get date/time stamp

                    string dateStamp = getStringBetweenString(fName, "", "_");
                    string timeStamp = getStringBetweenString(fName, "_", "_");

                    string processedDateTimeStamp = dateStamp + "_" + timeStamp;


                    //get employer name

                    string fNameMinusDate = fName.Substring(fName.IndexOf('_'));
                    string fNameMinusTime = getStringBetweenString(fNameMinusDate, "_", ".xlsm");

                    string employerName = getStringBetweenString(fNameMinusTime, "_", "_");


                    //search groups folder for employer name

                    try
                    {
                        string[] employerFolderNames = Directory.GetDirectories(GroupsFolderPath, employerName + "*", SearchOption.TopDirectoryOnly);

                        foreach (string employerFolder in employerFolderNames)
                        {
                            string groupName1 = Path.GetFileName(employerFolder);
                            string strToCompare = "-";

                            if (groupName1.Split().Count(r => r == strToCompare) > 1)
                            {
                                //do nothing and go to the next iteration
                            }
                            else
                            {
                                GroupName = groupName1;//delivers goupName without carriers
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.ToString());

                    }


                    //search for date/time stamp

                    string[] filesInEmployerFolder = Directory.GetFiles(GroupsFolderPath + GroupName + @"\Import\Done", processedDateTimeStamp + "*", SearchOption.AllDirectories);
                    foreach (string filePath in filesInEmployerFolder)
                    {
                        //MessageBox.Show(filePath);
                        string file = Path.GetFileName(filePath);
                        textBox3.Text = file;
                        Fname = file;


                        string destinationPath = @"C:\Users\14025\Documents\File Consultants\Brandon\Temp\" + Fname;
                        //read file in RichTextbox and Copy/Paste the file in question
                        if (!File.Exists(destinationPath))
                        {
                            File.Copy(filePath, destinationPath);
                            string fileText = File.ReadAllText(destinationPath);
                            richTextBox1.Text = fileText;


                        }
                        else
                        {
                            MessageBox.Show("This file already exists in the Temp folder.", "Can't Do It. Wouldn't be Proper.", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);

                            string fileText = File.ReadAllText(destinationPath);
                            richTextBox1.Text = fileText;
                        }
                    }
                    string origFileNameWithPFPFM = Fname.Substring(Fname.IndexOf('P'));



                    textBox5.Text = origFileNameWithPFPFM;

                    string erid = getStringBetweenString(origFileNameWithPFPFM, "M.", "_");





                    textBox1.Text = erid;
                    ERID = erid;

                    //get group name (without carriers) from erid

                    GetGroupName getGroupName = new GetGroupName();
                    getGroupName.GetGroupNameMethod(ERID);
                    string groupName = GetGroupName.GroupName;

                    this.Text = "FC Helper - " + groupName;

                    //get client ID
                    //GetClientID getClientID = new GetClientID();
                    //getClientID.GetClientIDMethod(groupName);


                    //ClientID = GetClientID.ClientID;

                    //textBox2.Text = ClientID;
                    textBox6.Text = processedDateTimeStamp;


                    //get group folder file path
                    //string groupFolderFilePath = @"C:\Users\14025\Documents\File Consultants\Groups\" + groupName;
                    //string[] groupFolderFilePathOfFile = Directory.GetFiles(groupFolderFilePath, Fname, SearchOption.AllDirectories);

                   

                    listView1.Items.Add(textBox5.Text);




                    //search for the groups error folder error reports
                    GetErrorReportFileName name = new GetErrorReportFileName();
                    name.GetErrorReportFileNameMethod(ERID, processedDateTimeStamp);

                    //search for the groups error folder for psvs
                    GetInputPSVFileNames names = new GetInputPSVFileNames();
                    names.GetInputPSVFileNamesMethod(ERID, processedDateTimeStamp);

                    //copy error reports to temp folder

                    foreach (string groupFolderItem in GetErrorReportFileName.ErrorReportGroupFolderLocations)
                    {

                        //get source path+filenames
                        string sourceFile = groupFolderItem;


                        string errorReportfilename = Path.GetFileName(sourceFile);
                        string destinationFile = @"C:\Users\14025\Documents\File Consultants\Brandon\Temp\" + errorReportfilename;

                        if (!File.Exists(destinationFile))
                        {
                            File.Copy(sourceFile, destinationFile);

                        }

                    }

                    //copy psv's to temp folder
                    foreach (string groupFolderItem in GetInputPSVFileNames.PSVFolderLocations)
                    {


                        //get source path+filenames
                        string sourceFile = groupFolderItem;


                        string psvfilename = Path.GetFileName(sourceFile);
                        string destinationFile = @"C:\Users\14025\Documents\File Consultants\Brandon\Temp\" + psvfilename;

                        if (!File.Exists(destinationFile))
                        {
                            File.Copy(sourceFile, destinationFile);

                        }
                    }

                    //add error report to listview2

                    foreach (string file in GetErrorReportFileName.ErrorReportFileNamesOnly)
                    {
                        listView2.Items.Add(file);

                    }
                    //add inputpsvs to listview3

                    foreach (string file in GetInputPSVFileNames.InputPSVFileNamesOnly)
                    {

                        listView3.Items.Add(file);

                    }
                    //add outputpsvs to listview4

                    foreach (string file in GetInputPSVFileNames.OutputPSVFileNamesOnly)
                    {
                        //MessageBox.Show(file);

                        listView4.Items.Add(file);

                    }

                    textBox1.Focus();
                    textBox1.SelectAll();
                    textBox1.Copy();
                }
            }
            //Fname = textBox3.Text;
            
        }//textbox 3 enter



        private void Form_Main_DragEnter(object sender, DragEventArgs e)
        {
            e.Effect = DragDropEffects.All;
        }

        private void Form_Main_FormClosing(object sender, FormClosingEventArgs e)
        {
            foreach (Control control in this.Controls)
            {
                if (control is TextBox)
                {
                    TextBox txtbox = (TextBox)control;
                    txtbox.Text = string.Empty;
                }
                else if (control is RichTextBox)
                {
                    RichTextBox richTextBox = (RichTextBox)control;
                    richTextBox.Text = string.Empty;
                }
                else if (control is ListView)
                {
                    ListView list = (ListView)control;
                    list.Items.Clear();
                }

            }

            //delete temp files

            string directory = @"C:\Users\14025\Documents\File Consultants\Brandon\Temp\";


            if (Directory.Exists(directory))
            {
                System.IO.DirectoryInfo di = new DirectoryInfo(directory);

                foreach (FileInfo file in di.GetFiles())
                {
                    file.Delete();
                }


            }
        }


        private void listView1_DoubleClick(object sender, EventArgs e)
        {
            //MessageBox.Show(Fname);

            string fileNameExtension = Path.GetExtension(Fname);//file extension

            string fileNameUpToPFDotPFM = getStringBetweenString(Fname, "","PF.PFM");

            string fileNameAfterPFDotPFMWithoutExtension = getStringBetweenString(Fname, "M.", ".");

            string newFileName = fileNameUpToPFDotPFM + "PF.PFM." + fileNameAfterPFDotPFMWithoutExtension + ".txt";



            string sourceFile = @"C:\Users\14025\Documents\File Consultants\Brandon\Temp\" + Fname;
            string destinationFile = @"C:\Users\14025\Documents\File Consultants\Brandon\Temp\" + newFileName;

            if (fileNameExtension != ".txt")
            {
                File.Move(sourceFile, destinationFile);
                Process.Start(destinationFile);

            }
            else if (fileNameExtension == ".txt")
            {
                //MessageBox.Show(Fname);

                try
                {
                    Process.Start(sourceFile);

                }
                catch (FileNotFoundException ex)
                {
                    Process.Start(destinationFile);
                }

            }



        }

        private void listView2_DoubleClick(object sender, EventArgs e)
        {
            //get selected filename

            string selectedFile = listView2.SelectedItems[0].Text;
            string destinationPath = @"C:\Users\14025\Documents\File Consultants\Brandon\Temp\";


            //start file in temp folder

            try
            {
                Process.Start(destinationPath + selectedFile);
            }
            catch (Win32Exception ex)
            {
                MessageBox.Show("File does not exist.\r\r" + ex, "Nonexistent File", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            catch (FileNotFoundException ex)
            {
                MessageBox.Show("File does not exist.\r\r" + ex, "Nonexistent File", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            catch (ObjectDisposedException ex)
            {
                MessageBox.Show("File does not exist.\r\r" + ex, "Nonexistent File", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }

        private void listView3_DoubleClick(object sender, EventArgs e)
        {

            //get selected filename

            string selectedFile = listView3.SelectedItems[0].Text;
            string destinationPath = @"C:\Users\14025\Documents\File Consultants\Brandon\Temp\";


            //start file in temp folder

            try
            {
                Process.Start(destinationPath + selectedFile);
            }
            catch (Win32Exception ex)
            {
                MessageBox.Show("File does not exist.\r\r" + ex, "Nonexistent File", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            catch (FileNotFoundException ex)
            {
                MessageBox.Show("File does not exist.\r\r" + ex, "Nonexistent File", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            catch (ObjectDisposedException ex)
            {
                MessageBox.Show("File does not exist.\r\r" + ex, "Nonexistent File", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void listView4_DoubleClick(object sender, EventArgs e)
        {

            //get selected filename

            string selectedFile = listView4.SelectedItems[0].Text;
            string destinationPath = @"C:\Users\14025\Documents\File Consultants\Brandon\Temp\";


            //start file in temp folder

            try
            {
                Process.Start(destinationPath+ selectedFile);
            }
            catch (Win32Exception ex)
            {
                MessageBox.Show("File does not exist.\r\r" + ex, "Nonexistent File", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            catch (FileNotFoundException ex)
            {
                MessageBox.Show("File does not exist.\r\r" + ex, "Nonexistent File", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            catch (ObjectDisposedException ex)
            {
                MessageBox.Show("File does not exist.\r\r" + ex, "Nonexistent File", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void textBox3_Enter(object sender, EventArgs e)
        {
            TextBox1Enter = true;
        }


        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            string f = textBox1.Text;
            bool fHasSpace = f.Contains(" ");

            if (fHasSpace == true)
            {
                char[] charToTrim = { ' ' };
                string result = textBox1.Text.Trim(charToTrim);
                textBox1.Text = (result);
                textBox1.Select(0, textBox1.Text.Length);
            }
        }

     
    }
}
