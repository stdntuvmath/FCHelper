using System;
using System.Data;
using System.Windows.Forms;
using System.IO;
using System.Diagnostics;


namespace FCHelper_v001
{
    public partial class Form_Connected_Claims_Folder_Prompt : Form
    {

        //global variables to refer to get something from this class

        public static string employerID;
        public static string carrierName;
        public static string employerName;
        public static string carrierCode;
        public static string CompleteNCCFilePath;
        public static string aetnaCarrierName;
        public static bool payflexLayout;
        public static bool existingLayout;
        public static bool enhancedVerification;
        public static bool layoutProvided;
        public static string LayoutFileName;
        public static bool aetnaTRADLayout;
        public static bool aetnaRXLayout;
        public static bool aetnaHMOLayout;
        private static string GroupName;
        private static string TempFolderPath = @"C:\Users\14025\Documents\File Consultants\Brandon\Temp\";




        public string TrimLastCharacter(string str)
        {
            if (String.IsNullOrEmpty(str))
            {
                return str;
            }
            else
            {
                return str.TrimEnd(str[str.Length - 1]);
            }
        }



        public Form_Connected_Claims_Folder_Prompt()
        {
            InitializeComponent();
            aetnaTRADLayout = false;
            aetnaRXLayout = false;
            aetnaHMOLayout = false;

        }

        private DataTable claimsCarriers;

        private void Form_Connected_Claims_Folder_Prompt_Load(object sender, EventArgs e)
        {

            listBox1.DataSource = GetClaimsCarrierNames();
            listBox1.DisplayMember = "Carriers"; 



        }

        private DataTable GetClaimsCarrierNames()//datatable for combobox
        {
            claimsCarriers = new DataTable();

            claimsCarriers.Columns.Add("Carriers", typeof(string));

            claimsCarriers.Rows.Add("Aetna HMO");
            claimsCarriers.Rows.Add("Aetna RX");
            claimsCarriers.Rows.Add("Aetna RX CUSTOM");
            claimsCarriers.Rows.Add("Aetna TRAD");
            claimsCarriers.Rows.Add("Aetna (Standalone)");
            claimsCarriers.Rows.Add("ACS");
            claimsCarriers.Rows.Add("AG Dental");
            claimsCarriers.Rows.Add("Allegiance");
            claimsCarriers.Rows.Add("Ameritas");
            claimsCarriers.Rows.Add("Anthem");
            claimsCarriers.Rows.Add("Anthem CVS Partnership Clients");
            claimsCarriers.Rows.Add("Anthem Vision Multi-Client");
            claimsCarriers.Rows.Add("Anthem.CC");

            claimsCarriers.Rows.Add("BCBS");
            claimsCarriers.Rows.Add("BCBS of AL");
            claimsCarriers.Rows.Add("BCBS of AR");
            claimsCarriers.Rows.Add("BCBS of FL");
            claimsCarriers.Rows.Add("BCBS of GA");
            claimsCarriers.Rows.Add("BCBS of KC");
            claimsCarriers.Rows.Add("BCBS of MI");
            claimsCarriers.Rows.Add("BCBS of MN");
            claimsCarriers.Rows.Add("BCBS of NE");
            claimsCarriers.Rows.Add("HCSC");
            claimsCarriers.Rows.Add("BCBS of Rochester");
            claimsCarriers.Rows.Add("Benecard");

            claimsCarriers.Rows.Add("Caremark");
            claimsCarriers.Rows.Add("Cigna");
            claimsCarriers.Rows.Add("CVS Caremark");

            claimsCarriers.Rows.Add("Davis Vision");
            claimsCarriers.Rows.Add("Delta Dental");
            claimsCarriers.Rows.Add("Delta Dental of MN");
            claimsCarriers.Rows.Add("Delta Dental of WA");
            claimsCarriers.Rows.Add("Delta Dental WDS");

            claimsCarriers.Rows.Add("Express Scripts");
            claimsCarriers.Rows.Add("Eye Med");

            claimsCarriers.Rows.Add("Gardian Dental");

            claimsCarriers.Rows.Add("Health Partners");
            claimsCarriers.Rows.Add("Health Span");
            claimsCarriers.Rows.Add("Horizon BCBS");

            claimsCarriers.Rows.Add("Kaiser");

            claimsCarriers.Rows.Add("MEDCO/Express Scrips (ESI)");
            claimsCarriers.Rows.Add("Medica Med");
            claimsCarriers.Rows.Add("Medica Rx");
            claimsCarriers.Rows.Add("Meritain Health");
            claimsCarriers.Rows.Add("Metlife Dental");


            claimsCarriers.Rows.Add("Navitus");

            claimsCarriers.Rows.Add("Optium");
            claimsCarriers.Rows.Add("PCS-Caremark");
            claimsCarriers.Rows.Add("PharmAvail, Inc. of Georgia");

            claimsCarriers.Rows.Add("SAVE-RX");
            claimsCarriers.Rows.Add("ScriptCare-RX Vendor");
            claimsCarriers.Rows.Add("Solstice Benefits");
            claimsCarriers.Rows.Add("Spectera Vision");
            claimsCarriers.Rows.Add("Superior Vision");

            claimsCarriers.Rows.Add("The Health Plan of West Virginia");//added by Claire Xie - Thu 5/21/2020 9:49 AM

            claimsCarriers.Rows.Add("UCCI");
            claimsCarriers.Rows.Add("UPMC");
            claimsCarriers.Rows.Add("UMR");
            claimsCarriers.Rows.Add("UHC");


            claimsCarriers.Rows.Add("VSP");

            return claimsCarriers;
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            string f = textBox1.Text;
            bool fHasSpace = f.Contains(" ");

            if (fHasSpace == true)
            {
                char[] charToTrim = {' '};               
                string result = textBox1.Text.Trim(charToTrim);
                textBox1.Text = (result);
                textBox1.Select(0,textBox1.Text.Length);
            }
        }

        private string folderName;
        public string groupName
        {
            get { return folderName; }
        }

        public string GetGroupName(out string groupname)
        {
            string ERID = textBox1.Text;
            string carrierName = textBox2.Text;
           
            

            //search groups folder for ERID, if !Exist then catch

            
            DirectoryInfo groupsDirectory = new DirectoryInfo(@"C:\Users\14025\Documents\File Consultants\Groups");
            DirectoryInfo[] directoryNames = groupsDirectory.GetDirectories("*" + ERID + "*.*");//gets all the directories with the ERID in it

         

           int counter = 0;
           foreach (DirectoryInfo name in directoryNames)
           {
                string groupName = name.Name;
                if (groupName == "- "+carrierName)
                {
                    DialogResult stuff = MessageBox.Show("This employer does not exist yet. " +
                                     "Please create this employer in PFM first.",
                                     "Hold up there slick...", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                    if (stuff == DialogResult.OK)
                    {

                        this.Close();

                    }
                }
                else
                {
                   
                    groupname = groupName;//delivers goupName out of foreach
                    return groupName;//pulls groupName out of for loop 
                    
                }

                
                              
           }
           groupname = groupName;//delivers goupName out of method
           return groupName;//pulls groupName out of for loop

           
        }

        //private string[] GetDataGridView1Data()
        //{
        //    string data = string.Empty;

        //    foreach (DataGridViewRow row in dataGridView1.Rows)
        //    {
        //        data = (string)row.Cells[0].Value;
        //        string[] array = { data };
        //        return data;
        //    }
        //    return data;
        //}

        private void button1_Click(object sender, EventArgs e)//OK button
        {
            //get multiERID's
            if (dataGridView1 == null)
            {
                dataGridView1.Enabled = false;
            }
            else
            {
                //string nERID = GetDataGridView1Data();
                //MessageBox.Show(nERID);

            }

            if (listView2.Items.Count == 0)
            {
                layoutProvided = false;
            }
            else if(listView2.Items.Count > 0)
            {
                layoutProvided = true;

                GetStringBetweenString getString = new GetStringBetweenString();

                foreach (ListViewItem item in listView2.Items)
                {
                    int selectionIndex = item.Index;
                    string layoutNameWithCrapOnIt = listView2.Items[selectionIndex].ToString();
                    string layoutFileName = getString.GetStringBetweenStringMethod(layoutNameWithCrapOnIt,"{","}");
                    LayoutFileName = layoutFileName;
                }

            }


            if (textBox1.Text == "")
            {
                MessageBox.Show("Please enter an ERID","Eh-hem...",MessageBoxButtons.OK,MessageBoxIcon.Hand);
            }
            else if (textBox2.Text == "")
            {
                MessageBox.Show("Please enter a carrier name.", "Eh-hem...", MessageBoxButtons.OK, MessageBoxIcon.Hand);

            }

            //GetGroupName();
            string groupName;

            GetGroupName(out groupName);
            

            for (int i = listBox2.Items.Count-1; i >= 0; --i)
            {
                             
                    DialogResult dr = MessageBox.Show("Are you sure you want to create the folder(s):\r" + groupName + " - " + listBox2.Items[i].ToString() + "\rIn the Groups folder?", "", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Warning);

                if (dr == DialogResult.Yes)
                {
                    

                    Class_CreateCCFolders cCCFolders = new Class_CreateCCFolders();
                    string folderName = cCCFolders.CreateCCFoldersMethod(groupName, listBox2.Items[i].ToString());//createCCFoldersMethod goes here
                    if (listBox2.Items[i].ToString() == "Aetna TRAD")
                    {
                        string sourceFile = @"\\phx-fs-02.payflex.com\GDrive\DataServicesGroup\Personal\Brandon\Tasks and Tools\Claims Tools\Claims File Layouts\Aetna TRAD Claims Copybook Layout.xlsx";
                        string destinationFile = folderName+ @"\Aetna TRAD Claims Copybook Layout.xlsx";

                        if (!File.Exists(destinationFile))
                        {
                            File.Copy(sourceFile, destinationFile);
                        }

                        aetnaCarrierName = "Aetna TRAD Claims Copybook Layout.xlsx";
                        aetnaTRADLayout = true;
                    }
                    else if (listBox2.Items[i].ToString() == "Aetna RX")
                    {
                        string sourceFile = @"\\phx-fs-02.payflex.com\GDrive\DataServicesGroup\Personal\Brandon\Tasks and Tools\Claims Tools\Claims File Layouts\Aetna RX Payflex File Layout.xlsx";
                        string destinationFile = folderName+ @"\Aetna RX Payflex File Layout.xlsx";

                        if (!File.Exists(destinationFile))
                        {
                            File.Copy(sourceFile, destinationFile);
                        }
                        aetnaCarrierName = "Aetna RX Payflex File Layout.xlsx";
                        aetnaRXLayout = true;
                    }
                    else if (listBox2.Items[i].ToString() == "Aetna HMO")
                    {
                        string sourceFile = @"\\phx-fs-02.payflex.com\GDrive\DataServicesGroup\Personal\Brandon\Tasks and Tools\Claims Tools\Claims File Layouts\Aetna HMO Payflex File Layout.xlsx";
                        string destinationFile = folderName+ @"\Aetna HMO Payflex File Layout.xlsx";

                        if (!File.Exists(destinationFile))
                        {
                            File.Copy(sourceFile, destinationFile);
                        }
                        aetnaCarrierName = "Aetna HMO Payflex File Layout.xlsx";
                        aetnaHMOLayout = true;
                    }
                    else if (checkBox1.Checked == true)
                    {
                        string sourceFile = @"\\phx-fs-02.payflex.com\GDrive\DataServicesGroup\Personal\Brandon\Tasks and Tools\Claims Tools\Claims File Layouts\Carrier File Guide.pdf";
                        string destinationFile = folderName + @"\Carrier File Guide.pdf";

                        if (!File.Exists(destinationFile))
                        {
                            File.Copy(sourceFile, destinationFile);
                        }
                    }
                    else if (checkBox1.Checked == false && listView2.Items.Count > 0)
                    {
                        //MessageBox.Show(folderName);
                        string sourceFile = TempFolderPath+LayoutFileName;
                        string destinationFile = folderName +@"\"+ LayoutFileName;

                        if (!File.Exists(destinationFile))
                        {
                            File.Copy(sourceFile, destinationFile);
                        }

                        File.Delete(sourceFile);
                    }
                }
                else if (dr == DialogResult.No)
                {
                    this.Close();
                    MessageBox.Show("Quitter.", "Fine. Whatever.", MessageBoxButtons.OK, MessageBoxIcon.Hand);
                }


                employerName = groupName;

                employerID = textBox1.Text;

                carrierName = listBox2.Items[i].ToString();

                carrierCode = textBox3.Text;

                Form_CDS_ClaimsTicketDescription form = new Form_CDS_ClaimsTicketDescription();
                form.Show();


                


                
            }

            //Rename NCC and open employers folder and open temp folder

            string todaysDate = DateTime.Today.ToString("yyyyMMdd");

            string newNCCFileName = todaysDate + "_" + employerName + "_NCC";
            RenameFile rename = new RenameFile();

            GetStringBetweenString getString1 = new GetStringBetweenString();

            foreach (ListViewItem item in listView1.Items)
            {
                int selectionIndex = item.Index;
                string NCCNameWithCrapOnIt = listView1.Items[selectionIndex].ToString();
                string NCCFileName = getString1.GetStringBetweenStringMethod(NCCNameWithCrapOnIt, "{", "}");
                rename.RenameFileMethod(TempFolderPath+ NCCFileName, newNCCFileName);
            }



            GetGroupName nameAndER = new GetGroupName();
            nameAndER.GetGroupNameMethod(employerID);
            string NameAndERID = FCHelper_v001.GetGroupName.GroupName;
            Process.Start(@"\\phx-fs-02.payflex.com\GDrive\DataServicesGroup\Personal\Brandon\Brandon's Staging Folder");
            Process.Start(@"C:\Users\14025\Documents\File Consultants\Groups\" + NameAndERID + @"\DOCS");
            //if (Directory.Exists(@"C:\Users\14025\Documents\File Consultants\Groups\" + NameAndERID+@"\DOCS"))
            //{
            //    Process.Start(@"C:\Users\14025\Documents\File Consultants\Groups\" + NameAndERID + @"\DOCS");

            //}
            //else if (Directory.Exists(@"C:\Users\14025\Documents\File Consultants\Groups\" + NameAndERID + @"\Docs"))
            //{
            //    Process.Start(@"C:\Users\14025\Documents\File Consultants\Groups\" + NameAndERID + @"\Docs");

            //}
            //else if (Directory.Exists(@"C:\Users\14025\Documents\File Consultants\Groups\" + NameAndERID + @"\docs"))
            //{
            //    Process.Start(@"C:\Users\14025\Documents\File Consultants\Groups\" + NameAndERID + @"\docs");

            //}

            Process.Start(@"C:\Users\14025\Documents\File Consultants\Temp");

        }

        private void button6_Click(object sender, EventArgs e)//Process Multiple ER's
        {
            

            foreach (DataGridViewRow row in dataGridView1.Rows)
            {

               
                string employerID = row.Cells[0].Value.ToString();
                MessageBox.Show(employerID);


               

                if (listView2 == null)
                {
                    layoutProvided = false;
                }


                FCHelper_v001.GetGroupName getGroupName = new GetGroupName();
                getGroupName.GetGroupNameMethod(employerID);

                GroupName = FCHelper_v001.GetGroupName.GroupName;
               


                for (int i = listBox2.Items.Count - 1; i >= 0; --i)
                {

                    DialogResult dr = MessageBox.Show("Are you sure you want to create the folder(s):\r" + GroupName + " - " + listBox2.Items[i].ToString() + "\rIn the Groups folder?", "", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Warning);

                    if (dr == DialogResult.Yes)
                    {


                        Class_CreateCCFolders cCCFolders = new Class_CreateCCFolders();
                        string folderName = cCCFolders.CreateCCFoldersMethod(GroupName, listBox2.Items[i].ToString());//createCCFoldersMethod goes here
                        if (listBox2.Items[i].ToString() == "Aetna TRAD")
                        {
                            string sourceFile = @"\\phx-fs-02.payflex.com\GDrive\DataServicesGroup\Personal\Brandon\Tasks and Tools\Claims Tools\Claims File Layouts\Aetna TRAD Claims Copybook Layout.xlsx";
                            string destinationFile = folderName + @"\Aetna TRAD Claims Copybook Layout.xlsx";

                            if (!File.Exists(destinationFile))
                            {
                                File.Copy(sourceFile, destinationFile);
                            }

                            aetnaCarrierName = "Aetna TRAD Claims Copybook Layout.xlsx";
                            aetnaTRADLayout = true;
                        }
                        else if (listBox2.Items[i].ToString() == "Aetna RX")
                        {
                            string sourceFile = @"\\phx-fs-02.payflex.com\GDrive\DataServicesGroup\Personal\Brandon\Tasks and Tools\Claims Tools\Claims File Layouts\Aetna RX Payflex File Layout.xlsx";
                            string destinationFile = folderName + @"\Aetna RX Payflex File Layout.xlsx";

                            if (!File.Exists(destinationFile))
                            {
                                File.Copy(sourceFile, destinationFile);
                            }
                            aetnaCarrierName = "Aetna RX Payflex File Layout.xlsx";
                            aetnaRXLayout = true;
                        }
                        else if (listBox2.Items[i].ToString() == "Aetna HMO")
                        {
                            string sourceFile = @"\\phx-fs-02.payflex.com\GDrive\DataServicesGroup\Personal\Brandon\Tasks and Tools\Claims Tools\Claims File Layouts\Aetna HMO Payflex File Layout.xlsx";
                            string destinationFile = folderName + @"\Aetna HMO Payflex File Layout.xlsx";

                            if (!File.Exists(destinationFile))
                            {
                                File.Copy(sourceFile, destinationFile);
                            }
                            aetnaCarrierName = "Aetna HMO Payflex File Layout.xlsx";
                            aetnaHMOLayout = true;
                        }
                    }
                    else if (dr == DialogResult.No)
                    {
                        this.Close();
                        MessageBox.Show("Quitter.", "Fine. Whatever.", MessageBoxButtons.OK, MessageBoxIcon.Hand);
                    }


                    employerName = GroupName;

                    employerID = textBox1.Text;

                    carrierName = listBox2.Items[i].ToString();

                    carrierCode = textBox3.Text;

                    Form_CDS_ClaimsTicketDescription form = new Form_CDS_ClaimsTicketDescription();
                    form.Show();






                }

                //Rename NCC and open employers folder and open temp folder

                string todaysDate = DateTime.Today.ToString("yyyyMMdd");

                string newNCCPathAndFileName = todaysDate + "_" + employerName + "_NCC";
                RenameFile rename = new RenameFile();
                rename.RenameFileMethod(CompleteNCCFilePath, newNCCPathAndFileName);

                GetGroupName nameAndER = new GetGroupName();
                nameAndER.GetGroupNameMethod(textBox1.Text);
                string NameAndERID = FCHelper_v001.GetGroupName.GroupName;

                if (Directory.Exists(@"C:\Users\14025\Documents\File Consultants\Groups\" + NameAndERID + @"\DOCS"))
                {
                    Process.Start(@"C:\Users\14025\Documents\File Consultants\Groups\" + NameAndERID + @"\DOCS");

                }
                else if (Directory.Exists(@"C:\Users\14025\Documents\File Consultants\Groups\" + NameAndERID + @"\Docs"))
                {
                    Process.Start(@"C:\Users\14025\Documents\File Consultants\Groups\" + NameAndERID + @"\Docs");

                }
                else if (Directory.Exists(@"C:\Users\14025\Documents\File Consultants\Groups\" + NameAndERID + @"\docs"))
                {
                    Process.Start(@"C:\Users\14025\Documents\File Consultants\Groups\" + NameAndERID + @"\docs");

                }

                Process.Start(@"C:\Users\14025\Documents\File Consultants\Temp");





            }
        }

        private void button2_Click(object sender, EventArgs e)//Cancel button
        {
            
            this.Close();
        }

        private void button3_Click(object sender, EventArgs e)//Open Groups folder button
        {
            string groupFolderPath = @"C:\Users\14025\Documents\File Consultants\Groups";
            Process.Start(groupFolderPath);
        }

        private void button4_Click(object sender, EventArgs e)//Add Multiple Carriers
        {
            if (textBox2.Text == "")
            {
                MessageBox.Show("Please select a carrier to add.", "Eh-hem...", MessageBoxButtons.OK, MessageBoxIcon.Question);
            }
            else
            {
                listBox2.Items.Add(textBox2.Text);
            }
            
            
        }

        private void button5_Click(object sender, EventArgs e)//Remove Multiple Carriers
        {
            if (listBox2.SelectedIndex < 0)
            {
                MessageBox.Show("Please select a carrier to remove.","Eh-hem...",MessageBoxButtons.OK,MessageBoxIcon.Question);
            }
            else
            {
                listBox2.Items.RemoveAt(listBox2.SelectedIndex);
            }
            
        }



        private void textBox2_TextChanged(object sender, EventArgs e)//filters through combobox when you type
        {
            //DataView dvCarriers = claimsCarriers.DefaultView;

            //dvCarriers.RowFilter = "Carriers LIKE '%" + textBox2.Text + "%'";//SQL command - column LIKE '%text%'
        }


    

        private void listBox1_Click(object sender, EventArgs e)
        {
           
            textBox2.Text = listBox1.Text;


        }

     

   

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox1.Checked)
            {
                listView2.Enabled = false;
                listView2.BackColor = System.Drawing.Color.DarkGray;
                payflexLayout = true;
            }
            else
            {
                listView2.Enabled = true;
                listView2.BackColor = System.Drawing.Color.White;
                payflexLayout = false;

            }
        }

        private void checkBox2_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox2.Checked)
            {
                
                enhancedVerification = true;
            }
            else
            {
                
                enhancedVerification = false;

            }
        }
        private void checkBox3_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox3.Checked)
            {
                listView2.Enabled = false;
                listView2.BackColor = System.Drawing.Color.DarkGray;
                existingLayout = true;
            }
            else
            {
                listView2.Enabled = true;
                listView2.BackColor = System.Drawing.Color.White;
                existingLayout = false;

            }
        }

        private void listView1_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                e.Effect = DragDropEffects.Copy;//copy the file
            }
        }

        private void listView1_DragDrop(object sender, DragEventArgs e)
        {
            string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);//gets the file dropped

            int counter = 0;
            

            foreach (string completeNCCFilePath in files)
            {


                counter++;
                string sourcefilePath = System.IO.Path.GetDirectoryName(completeNCCFilePath);//complete path from drop
                string fileName = System.IO.Path.GetFileName(completeNCCFilePath);
                string destinationPath = @"C:\Users\14025\Documents\File Consultants\Brandon\Temp\" + fileName;


                //ImageList1.Images.Add(Icon.ExtractAssociatedIcon(fileName))
                //System.Drawing.Icon.ExtractAssociatedIcon(completeFilePath);


                if (!File.Exists(destinationPath))
                {
                    File.Copy(sourcefilePath+"\\"+fileName, destinationPath);
                    listView1.Items.Add(fileName);
                    CompleteNCCFilePath = destinationPath;
                }
                else
                {
                    MessageBox.Show("This file already exists in this this employers folder.", "Can't Do It. Wouldn't be Proper.", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                }

            }
        }

        private void listView2_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                e.Effect = DragDropEffects.Copy;//copy the file
            }
        }

        private void listView2_DragDrop(object sender, DragEventArgs e)
        {
            string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);

            int counter = 0;
            //ImageList ImageList1 = new ImageList ();
            //string[] filesInDir = Directory.GetFiles(@"C:\Users\14025\Documents\File Consultants\Brandon\" + textBox1.Text + "_" + textBox2.Text);

            //foreach (string file in filesInDir)
            //{

            //    Icon stuff = Icon.ExtractAssociatedIcon(file);
            //    ImageList1.Images.Add(stuff);

            //}

            foreach (string completeFileLayoutFilePath in files)
            {


                counter++;
                string filePath = System.IO.Path.GetDirectoryName(completeFileLayoutFilePath);
                string fileName = System.IO.Path.GetFileName(completeFileLayoutFilePath);
                string destinationPath = @"C:\Users\14025\Documents\File Consultants\Brandon\Temp\" + fileName;


                //ImageList1.Images.Add(Icon.ExtractAssociatedIcon(fileName))
                //System.Drawing.Icon.ExtractAssociatedIcon(completeFilePath);


                if (!File.Exists(destinationPath))
                {
                    File.Copy(completeFileLayoutFilePath, destinationPath);
                    listView2.Items.Add(fileName);
                    CompleteNCCFilePath = completeFileLayoutFilePath;
                }
                else
                {
                    MessageBox.Show("This file already exists in this this employers folder.", "Can't Do It. Wouldn't be Proper.", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                }

            }
        }

        private void Form_Connected_Claims_Folder_Prompt_FormClosed(object sender, FormClosedEventArgs e)
        {
            var selectedItems = listView1.SelectedItems;


            foreach (ListViewItem selectedItem in selectedItems)
            {
                listView1.Items.Remove(selectedItem);

                string stuff = selectedItem.Text;

                string fileToDelete = @"C:\Users\14025\Documents\File Consultants\Brandon\File Consultants\Brandon\Temp\" + stuff;


                if (File.Exists(fileToDelete))
                {
                    try
                    {
                        File.Delete(fileToDelete);
                    }
                    catch (System.Exception ex)
                    {
                        MessageBox.Show(ex.StackTrace);
                    }

                }

            }
        }

        private void listView1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Delete)
            {
                var selectedItems = listView1.SelectedItems;


                foreach (ListViewItem selectedItem in selectedItems)
                {
                    listView1.Items.Remove(selectedItem);

                    string stuff = selectedItem.Text;

                    string fileToDelete = @"C:\Users\14025\Documents\File Consultants\Brandon\File ConsultantsTemp\" + stuff;


                    if (File.Exists(fileToDelete))
                    {
                        try
                        {
                            File.Delete(fileToDelete);
                        }
                        catch (System.Exception ex)
                        {
                            MessageBox.Show(ex.StackTrace);
                        }

                    }

                }
                e.Handled = true;
            }
        }

        private void listView2_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Delete)
            {
                var selectedItems = listView2.SelectedItems;


                foreach (ListViewItem selectedItem in selectedItems)
                {
                    listView2.Items.Remove(selectedItem);

                    string stuff = selectedItem.Text;

                    string fileToDelete = @"C:\Users\14025\Documents\File Consultants\Brandon\File Consultants\Temp" + stuff;


                    if (File.Exists(fileToDelete))
                    {
                        try
                        {
                            File.Delete(fileToDelete);
                        }
                        catch (System.Exception ex)
                        {
                            MessageBox.Show(ex.StackTrace);
                        }

                    }

                }
                e.Handled = true;
            }
        }

       
    }
}
