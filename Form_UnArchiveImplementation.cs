using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;

namespace FCHelper_v001
{

    public partial class Form_UnArchiveImplementation : Form
    {

        private static string ZippedFileName;
        private static string ZippedFilePath;
        private static string FileNameWithoutExtension;
        private static string RtfFileToTextFile;
        private static bool CSV_isChecked = false;

        public static string[] ImpDataArray;





        public Form_UnArchiveImplementation()
        {
            InitializeComponent();

            textBox1.TabIndex = 0;
            button1.TabIndex = 1;
            button2.TabIndex = 2;

        }

        private void Form_UnArchiveImplementation_Load(object sender, EventArgs e)
        {
            textBox1.Select();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string erid1 = textBox1.Text;

            GetStringBetweenString getString = new GetStringBetweenString();
            GetGroupName getGroupName = new GetGroupName();
            getGroupName.GetGroupNameMethod(erid1);
            string groupName = GetGroupName.GroupName;

            //search the archive folder for the erid
            string zippedCSVfiles = @"C:\Users\14025\Documents\File Consultants\Brandon\Archive\Backup_CSV\";

            string archivedFilesPath = @"C:\Users\14025\Documents\File Consultants\Brandon\Archive";
            string[] archivedFiles = Directory.GetFiles(archivedFilesPath);
            string[] achivedCSVFiles = Directory.GetFiles(zippedCSVfiles);

            if (CSV_isChecked == true)
            {
                foreach (string zippedFilePath in achivedCSVFiles)
                {
                    if (zippedFilePath.Contains(erid1))
                    {
                        ZippedFilePath = zippedFilePath;

                        string zippedFileName = Path.GetFileName(zippedFilePath);
                        //MessageBox.Show(zippedFileName);
                        ZippedFileName = zippedFileName;
                        FileNameWithoutExtension = getString.GetStringBetweenStringMethod(zippedFileName, "", ".");
                    }
                    else
                    {
                        //MessageBox.Show("The employer: "+groupName+" is not an archived implementation.","Employer is not Archived.",MessageBoxButtons.OK,MessageBoxIcon.Warning);
                    }
                }


                string pathOfFileToExtract1 = ZippedFilePath;
                //string groupNameWithoutSpaces = groupName.Replace(" ", "");
                string groupNameWithUnderscore1 = groupName.Replace(" - ", "_");
                string destinationFolderName1 = @"C:\Users\14025\Documents\File Consultants\Brandon\" + groupNameWithUnderscore1;



                DialogResult result1 = MessageBox.Show("Are you sure you want to resurrect this Implementation from the depths of Hades?\r\r" + groupName, "Resurrect this Implementation?", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);

                if (result1 == DialogResult.Yes)
                {





                    //move contents of folder to new employer folder in Brandon Folder

                    if (!Directory.Exists(groupName))
                    {

                        Directory.CreateDirectory(destinationFolderName1);


                        try
                        {

                            System.IO.Compression.ZipFile.ExtractToDirectory(pathOfFileToExtract1, destinationFolderName1);


                        }
                        catch (System.IO.IOException ex)
                        {
                            //MessageBox.Show("Method: ZipFileMethod()\r\r" + ex, "Could not unzip/archive employer directory", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                        catch (System.ArgumentException ex)
                        {
                            // MessageBox.Show("Method: ZipFileMethod()\r\r" + ex, "Could not zip/archive employer directory", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                        catch (System.NotSupportedException ex)
                        {
                            // MessageBox.Show("Method: ZipFileMethod()\r\r" + ex, "Could not zip/archive employer directory", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }




                    }
                    else
                    {

                    }

                    //MessageBox.Show(destinationFolderName);
                    //get and read the rtf file
                    string[] unArchivedFiles = Directory.GetFiles(destinationFolderName1);
                    string sourceFileName = @"C:\Users\14025\Documents\File Consultants\Brandon\" + groupNameWithUnderscore1 + @"\" + groupNameWithUnderscore1 + @"_Notes.rtf";
                    string sourceFileNameTXT = @"C:\Users\14025\Documents\File Consultants\Brandon\" + groupNameWithUnderscore1 + @"\" + groupNameWithUnderscore1 + @"_Notes.txt";

                    RtfFileToTextFile = sourceFileNameTXT;
                    string notesFileName = @"C:\Users\14025\Documents\File Consultants\Brandon\Notes\" + groupNameWithUnderscore1 + @"_Notes.rtf";


                    foreach (string unArchivedFile in unArchivedFiles)
                    {


                        if (unArchivedFile.Contains(".rtf"))
                        {
                            //create a copy in .txt format, inside the unarchived folder
                            if (!File.Exists(sourceFileNameTXT))
                            {
                                try
                                {
                                    File.Copy(sourceFileName, sourceFileNameTXT);

                                }
                                catch (IOException ex)
                                {
                                    MessageBox.Show("File could not be copied because of the following error: \r\r" + ex, "File Could Not Be Copied.", MessageBoxButtons.OK, MessageBoxIcon.Error);

                                }
                            }
                            else
                            {
                                MessageBox.Show("File already exists in the Notes folder", "File Could Not Be Moved.", MessageBoxButtons.OK, MessageBoxIcon.Error);

                            }


                            //MessageBox.Show(unArchivedFile);

                            //move rtf file to the Notes folder

                            if (!File.Exists(notesFileName))
                            {
                                try
                                {
                                    File.Move(sourceFileName, notesFileName);

                                }
                                catch (IOException ex)
                                {
                                    MessageBox.Show("File could not be moved because of the following error: \r\r" + ex, "File Could Not Be Moved.", MessageBoxButtons.OK, MessageBoxIcon.Error);

                                }
                            }
                            else
                            {
                                MessageBox.Show("File already exists in the Notes folder", "File Could Not Be Moved.", MessageBoxButtons.OK, MessageBoxIcon.Error);

                            }
                        }
                    }


                    //get Implemenation Details from the text file



                    if (File.Exists(sourceFileNameTXT))
                    {
                        string allTextFileData = File.ReadAllText(sourceFileNameTXT);

                        if (allTextFileData.Contains("Details:"))
                        {
                            //MessageBox.Show(allTextFileData);

                            string isolateImpData = getString.GetStringBetweenStringMethod(allTextFileData, "Details:", "}");

                            MessageBox.Show(isolateImpData);



                            string impData = isolateImpData.Substring(12);
                            MessageBox.Show(impData);

                            //parse and store impData into a string array

                            string[] impDataArray = impData.Split(',');
                            ImpDataArray = impDataArray;




                        }
                    }
                    else
                    {
                        MessageBox.Show("File does not exist in folder", "Cannot Open File", MessageBoxButtons.OK, MessageBoxIcon.Error);

                    }



                    //add information back to Excel database

                    string ername = ImpDataArray[0];
                    string erid2 = ImpDataArray[1];
                    string region = ImpDataArray[2];
                    string segment = ImpDataArray[3];
                    string effDate = ImpDataArray[4];
                    string curProd = ImpDataArray[5];
                    string addProd = ImpDataArray[6];
                    string newImp = ImpDataArray[7];
                    string AM_IM = ImpDataArray[8];
                    string impDdline = ImpDataArray[9];
                    string sftpFlag = ImpDataArray[10];

                    string inConName = ImpDataArray[11];
                    string inConPhone = ImpDataArray[12];
                    string inConEmail = ImpDataArray[13];
                    string inConType = ImpDataArray[14];

                    string exConName = ImpDataArray[15];
                    string exConPhone = ImpDataArray[16];
                    string exConEmail = ImpDataArray[17];
                    string exConType = ImpDataArray[18];
                    string fileType = ImpDataArray[19];

                    string chkbx1 = ImpDataArray[20];
                    string chkbx2 = ImpDataArray[21];
                    string chkbx3 = ImpDataArray[22];
                    string chkbx4 = ImpDataArray[23];
                    string chkbx5 = ImpDataArray[24];
                    string chkbx6 = ImpDataArray[25];
                    string chkbx7 = ImpDataArray[26];
                    string chkbx8 = ImpDataArray[27];
                    string chkbx9 = ImpDataArray[28];
                    string chkbx10 = ImpDataArray[29];

                    string inConName2 = ImpDataArray[30];
                    string inConPhone2 = ImpDataArray[31];
                    string inConEmail2 = ImpDataArray[32];
                    string inConType2 = ImpDataArray[33];

                    string inConName3 = ImpDataArray[34];
                    string inConPhone3 = ImpDataArray[35];
                    string inConEmail3 = ImpDataArray[36];
                    string inConType3 = ImpDataArray[37];

                    string inConName4 = ImpDataArray[38];
                    string inConPhone4 = ImpDataArray[39];
                    string inConEmail4 = ImpDataArray[40];
                    string inConType4 = ImpDataArray[41];

                    string exConName2 = ImpDataArray[42];
                    string exConPhone2 = ImpDataArray[43];
                    string exConEmail2 = ImpDataArray[44];
                    string exConType2 = ImpDataArray[45];

                    string exConName3 = ImpDataArray[46];
                    string exConPhone3 = ImpDataArray[47];
                    string exConEmail3 = ImpDataArray[48];
                    string exConType3 = ImpDataArray[49];

                    string exConName4 = ImpDataArray[50];
                    string exConPhone4 = ImpDataArray[51];
                    string exConEmail4 = ImpDataArray[52];
                    string exConType4 = ImpDataArray[53];



                    ExcelDataBasePush db = new ExcelDataBasePush();
                    db.ExcelDataBasePushMethod(ername, erid2, region, segment,
                                            effDate, curProd, addProd, newImp,
                                            AM_IM, impDdline, sftpFlag,
                                            inConName, inConPhone, inConEmail, inConType,
                                            exConName, exConPhone, exConEmail, exConType, fileType,
                                            inConName2, inConPhone2, inConEmail2, inConType2,
                                            inConName3, inConPhone3, inConEmail3, inConType3,
                                            inConName4, inConPhone4, inConEmail4, inConType4,
                                            exConName2, exConPhone2, exConEmail2, exConType2,
                                            exConName3, exConPhone3, exConEmail3, exConType3,
                                            exConName4, exConPhone4, exConEmail4, exConType4,
                                            chkbx1, chkbx2, chkbx3, chkbx4, chkbx5, chkbx6,
                                            chkbx7, chkbx8, chkbx9, chkbx10);


                }
                else if (result1 == DialogResult.No)
                {

                }

                if (File.Exists(RtfFileToTextFile))
                {
                    File.Delete(RtfFileToTextFile);
                }
                else
                {
                    MessageBox.Show("Cannot delete .txt file. File does not exist in folder", "Cannot Delete File", MessageBoxButtons.OK, MessageBoxIcon.Error);

                }

                //delete completed .zip file
                if (File.Exists(ZippedFilePath))
                {
                    File.Delete(ZippedFilePath);
                }
                else
                {
                    MessageBox.Show("Cannot delete .zip file. File does not exist in Archive folder", "Cannot Delete File", MessageBoxButtons.OK, MessageBoxIcon.Error);

                }

                // refresh Implementation List
                Form_OpenImplementationList form1 = new Form_OpenImplementationList();
                form1.Show();
            }
            else
            {
                foreach (string zippedFilePath in archivedFiles)
                {
                    if (zippedFilePath.Contains(erid1))
                    {
                        ZippedFilePath = zippedFilePath;

                        string zippedFileName = Path.GetFileName(zippedFilePath);
                        //MessageBox.Show(zippedFileName);
                        ZippedFileName = zippedFileName;
                        FileNameWithoutExtension = getString.GetStringBetweenStringMethod(zippedFileName, "", ".");
                    }
                    else
                    {
                        //MessageBox.Show("The employer: "+groupName+" is not an archived implementation.","Employer is not Archived.",MessageBoxButtons.OK,MessageBoxIcon.Warning);
                    }
                }


                string pathOfFileToExtract = ZippedFilePath;
                //string groupNameWithoutSpaces = groupName.Replace(" ", "");
                string groupNameWithUnderscore = groupName.Replace(" - ", "_");
                string destinationFolderName = @"C:\Users\14025\Documents\File Consultants\Brandon\" + groupNameWithUnderscore;



                DialogResult result = MessageBox.Show("Are you sure you want to resurrect this Implementation from the depths of Hades?\r\r" + groupName, "Resurrect this Implementation?", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);

                if (result == DialogResult.Yes)
                {





                    //move contents of folder to new employer folder in Brandon Folder

                    if (!Directory.Exists(groupName))
                    {

                        Directory.CreateDirectory(destinationFolderName);


                        try
                        {

                            System.IO.Compression.ZipFile.ExtractToDirectory(pathOfFileToExtract, destinationFolderName);


                        }
                        catch (System.IO.IOException ex)
                        {
                            //MessageBox.Show("Method: ZipFileMethod()\r\r" + ex, "Could not unzip/archive employer directory", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                        catch (System.ArgumentException ex)
                        {
                            // MessageBox.Show("Method: ZipFileMethod()\r\r" + ex, "Could not zip/archive employer directory", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                        catch (System.NotSupportedException ex)
                        {
                            // MessageBox.Show("Method: ZipFileMethod()\r\r" + ex, "Could not zip/archive employer directory", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }




                    }
                    else
                    {

                    }

                    //MessageBox.Show(destinationFolderName);
                    //get and read the rtf file
                    string[] unArchivedFiles = Directory.GetFiles(destinationFolderName);
                    string sourceFileName = @"C:\Users\14025\Documents\File Consultants\Brandon\" + groupNameWithUnderscore + @"\" + groupNameWithUnderscore + @"_Notes.rtf";
                    string sourceFileNameTXT = @"C:\Users\14025\Documents\File Consultants\Brandon\" + groupNameWithUnderscore + @"\" + groupNameWithUnderscore + @"_Notes.txt";

                    RtfFileToTextFile = sourceFileNameTXT;
                    string notesFileName = @"C:\Users\14025\Documents\File Consultants\Brandon\Notes\" + groupNameWithUnderscore + @"_Notes.rtf";


                    foreach (string unArchivedFile in unArchivedFiles)
                    {


                        if (unArchivedFile.Contains(".rtf"))
                        {
                            //create a copy in .txt format, inside the unarchived folder
                            if (!File.Exists(sourceFileNameTXT))
                            {
                                try
                                {
                                    File.Copy(sourceFileName, sourceFileNameTXT);

                                }
                                catch (IOException ex)
                                {
                                    MessageBox.Show("File could not be copied because of the following error: \r\r" + ex, "File Could Not Be Copied.", MessageBoxButtons.OK, MessageBoxIcon.Error);

                                }
                            }
                            else
                            {
                                MessageBox.Show("File already exists in the Notes folder", "File Could Not Be Moved.", MessageBoxButtons.OK, MessageBoxIcon.Error);

                            }


                            //MessageBox.Show(unArchivedFile);

                            //move rtf file to the Notes folder

                            if (!File.Exists(notesFileName))
                            {
                                try
                                {
                                    File.Move(sourceFileName, notesFileName);

                                }
                                catch (IOException ex)
                                {
                                    MessageBox.Show("File could not be moved because of the following error: \r\r" + ex, "File Could Not Be Moved.", MessageBoxButtons.OK, MessageBoxIcon.Error);

                                }
                            }
                            else
                            {
                                MessageBox.Show("File already exists in the Notes folder", "File Could Not Be Moved.", MessageBoxButtons.OK, MessageBoxIcon.Error);

                            }
                        }
                    }


                    //get Implemenation Details from the text file



                    if (File.Exists(sourceFileNameTXT))
                    {
                        string allTextFileData = File.ReadAllText(sourceFileNameTXT);

                        if (allTextFileData.Contains("Details:"))
                        {
                            //MessageBox.Show(allTextFileData);

                            string isolateImpData = getString.GetStringBetweenStringMethod(allTextFileData, "Details:", "}");

                            MessageBox.Show(isolateImpData);



                            string impData = isolateImpData.Substring(12);
                            MessageBox.Show(impData);

                            //parse and store impData into a string array

                            string[] impDataArray = impData.Split('|');
                            ImpDataArray = impDataArray;




                        }
                    }
                    else
                    {
                        MessageBox.Show("File does not exist in folder", "Cannot Open File", MessageBoxButtons.OK, MessageBoxIcon.Error);

                    }



                    //add information back to Excel database

                    string ername = ImpDataArray[0];
                    string erid2 = ImpDataArray[1];
                    string region = ImpDataArray[2];
                    string segment = ImpDataArray[3];
                    string effDate = ImpDataArray[4];
                    string curProd = ImpDataArray[5];
                    string addProd = ImpDataArray[6];
                    string newImp = ImpDataArray[7];
                    string AM_IM = ImpDataArray[8];
                    string impDdline = ImpDataArray[9];
                    string sftpFlag = ImpDataArray[10];

                    string inConName = ImpDataArray[11];
                    string inConPhone = ImpDataArray[12];
                    string inConEmail = ImpDataArray[13];
                    string inConType = ImpDataArray[14];

                    string exConName = ImpDataArray[15];
                    string exConPhone = ImpDataArray[16];
                    string exConEmail = ImpDataArray[17];
                    string exConType = ImpDataArray[18];
                    string fileType = ImpDataArray[19];

                    string chkbx1 = ImpDataArray[20];
                    string chkbx2 = ImpDataArray[21];
                    string chkbx3 = ImpDataArray[22];
                    string chkbx4 = ImpDataArray[23];
                    string chkbx5 = ImpDataArray[24];
                    string chkbx6 = ImpDataArray[25];
                    string chkbx7 = ImpDataArray[26];
                    string chkbx8 = ImpDataArray[27];
                    string chkbx9 = ImpDataArray[28];
                    string chkbx10 = ImpDataArray[29];

                    string inConName2 = ImpDataArray[30];
                    string inConPhone2 = ImpDataArray[31];
                    string inConEmail2 = ImpDataArray[32];
                    string inConType2 = ImpDataArray[33];

                    string inConName3 = ImpDataArray[34];
                    string inConPhone3 = ImpDataArray[35];
                    string inConEmail3 = ImpDataArray[36];
                    string inConType3 = ImpDataArray[37];

                    string inConName4 = ImpDataArray[38];
                    string inConPhone4 = ImpDataArray[39];
                    string inConEmail4 = ImpDataArray[40];
                    string inConType4 = ImpDataArray[41];

                    string exConName2 = ImpDataArray[42];
                    string exConPhone2 = ImpDataArray[43];
                    string exConEmail2 = ImpDataArray[44];
                    string exConType2 = ImpDataArray[45];

                    string exConName3 = ImpDataArray[46];
                    string exConPhone3 = ImpDataArray[47];
                    string exConEmail3 = ImpDataArray[48];
                    string exConType3 = ImpDataArray[49];

                    string exConName4 = ImpDataArray[50];
                    string exConPhone4 = ImpDataArray[51];
                    string exConEmail4 = ImpDataArray[52];
                    string exConType4 = ImpDataArray[53];



                    ExcelDataBasePush db = new ExcelDataBasePush();
                    db.ExcelDataBasePushMethod(ername, erid2, region, segment,
                                            effDate, curProd, addProd, newImp,
                                            AM_IM, impDdline, sftpFlag,
                                            inConName, inConPhone, inConEmail, inConType,
                                            exConName, exConPhone, exConEmail, exConType, fileType,
                                            inConName2, inConPhone2, inConEmail2, inConType2,
                                            inConName3, inConPhone3, inConEmail3, inConType3,
                                            inConName4, inConPhone4, inConEmail4, inConType4,
                                            exConName2, exConPhone2, exConEmail2, exConType2,
                                            exConName3, exConPhone3, exConEmail3, exConType3,
                                            exConName4, exConPhone4, exConEmail4, exConType4,
                                            chkbx1, chkbx2, chkbx3, chkbx4, chkbx5, chkbx6,
                                            chkbx7, chkbx8, chkbx9, chkbx10);


                }
                else if (result == DialogResult.No)
                {

                }

                if (File.Exists(RtfFileToTextFile))
                {
                    File.Delete(RtfFileToTextFile);
                }
                else
                {
                    MessageBox.Show("Cannot delete .txt file. File does not exist in folder", "Cannot Delete File", MessageBoxButtons.OK, MessageBoxIcon.Error);

                }

                //delete completed .zip file
                if (File.Exists(ZippedFilePath))
                {
                    File.Delete(ZippedFilePath);
                }
                else
                {
                    MessageBox.Show("Cannot delete .zip file. File does not exist in Archive folder", "Cannot Delete File", MessageBoxButtons.OK, MessageBoxIcon.Error);

                }

                // refresh Implementation List
                Form_OpenImplementationList form = new Form_OpenImplementationList();
                form.Show();
            }


            

        }


        private void button2_Click(object sender, EventArgs e)//Cancel
        {
            this.Dispose();
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox1.Checked == true)
            {
                CSV_isChecked = true;
            }
            else
            {
                CSV_isChecked = false;
            }
        }
    }
}
