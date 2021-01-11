using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace FCHelper_v001
{
    public partial class Form_CDS_ClaimsTicketDescription : Form
    {
        public Form_CDS_ClaimsTicketDescription()
        {
            InitializeComponent();
        }

        private void Form_CDS_ClaimsTicketDescription_Load(object sender, EventArgs e)
        {

            string ername = Form_Connected_Claims_Folder_Prompt.employerName;
            string erid = Form_Connected_Claims_Folder_Prompt.employerID;
            string carriername = Form_Connected_Claims_Folder_Prompt.carrierName;            
            string subjectLine = ername + " - " + carriername;
            string backOfficeCode = Form_Connected_Claims_Folder_Prompt.carrierCode;
            string AetnaCarrierName = Form_Connected_Claims_Folder_Prompt.aetnaCarrierName;
            string LayoutFileName = Form_Connected_Claims_Folder_Prompt.LayoutFileName;
            bool payFlexLayout = Form_Connected_Claims_Folder_Prompt.payflexLayout;
            bool existingLayout = Form_Connected_Claims_Folder_Prompt.existingLayout;
            bool EV = Form_Connected_Claims_Folder_Prompt.enhancedVerification;
            bool layoutPrvded = Form_Connected_Claims_Folder_Prompt.layoutProvided;
            bool AetnaTRADLayout = Form_Connected_Claims_Folder_Prompt.aetnaTRADLayout;
            bool AetnaRXLayout = Form_Connected_Claims_Folder_Prompt.aetnaRXLayout;
            bool AetnaHMOLayout = Form_Connected_Claims_Folder_Prompt.aetnaHMOLayout;






            textBox1.Text = subjectLine;

            if (AetnaTRADLayout == true && EV == true)
            {
                string eVee = "EV has been setup in BackOffice";
                string layoutGd = AetnaCarrierName;
                string noLayoutGd = "Layout Not Submitted Waiting For Reply From Carrier";
                richTextBox1.Text = "Connected Claims setup for " + subjectLine + "\n" +
                                    @"Group Folder: C:\Users\14025\Documents\File Consultants\Groups\" + subjectLine + @"\" + "\n" +
                                    "Back Office: Carrier code " + "AET_TRAD" + " is setup.\n" +
                                    "File Layout Guide: " + layoutGd + "\n" +
                                    "NCC: Has been stored in this groups employer DOCS folder.\n" +
                                    "EV: " + eVee + "\n";
            }
            else if (AetnaTRADLayout == true && EV == false)
            {
                string eVee = "EV hasn't been setup in BackOffice";
                string layoutGd = AetnaCarrierName;
                string noLayoutGd = "Layout Not Submitted Waiting For Reply From Carrier";
                richTextBox1.Text = "Connected Claims setup for " + subjectLine + "\n" +
                                    @"Group Folder: C:\Users\14025\Documents\File Consultants\Groups\" + subjectLine + @"\" + "\n" +
                                    "Back Office: Carrier code " + "AET_TRAD" + " is setup.\n" +
                                    "File Layout Guide: " + layoutGd + "\n" +
                                    "NCC: Has been stored in this groups employer DOCS folder.\n" +
                                    "EV: " + eVee + "\n" +
                                    "Special Notes: No Debit Card Client";
            }
            
            else if (AetnaRXLayout == true && EV == true)
            {
                string eVee = "EV has been setup in BackOffice";
                string layoutGd = AetnaCarrierName;
                string noLayoutGd = "Layout Not Submitted Waiting For Reply From Carrier";
                richTextBox1.Text = "Connected Claims setup for " + subjectLine + "\n" +
                                    @"Group Folder: C:\Users\14025\Documents\File Consultants\Groups\" + subjectLine + @"\" + "\n" +
                                    "Back Office: Carrier code " + "AET_RX" + " is setup.\n" +
                                    "File Layout Guide: " + layoutGd + "\n" +
                                    "NCC: Has been stored in this groups employer DOCS folder.\n" +
                                    "EV: " + eVee + "\n";
            }
            else if (AetnaRXLayout == true && EV == false)
            {
                string eVee = "EV hasn't been setup in BackOffice";
                string layoutGd = AetnaCarrierName;
                string noLayoutGd = "Layout Not Submitted Waiting For Reply From Carrier";
                richTextBox1.Text = "Connected Claims setup for " + subjectLine + "\n" +
                                    @"Group Folder: C:\Users\14025\Documents\File Consultants\Groups\" + subjectLine + @"\" + "\n" +
                                    "Back Office: Carrier code " + "AET_RX" + " is setup.\n" +
                                    "File Layout Guide: " + layoutGd + "\n" +
                                    "NCC: Has been stored in this groups employer DOCS folder.\n" +
                                    "EV: " + eVee + "\n" +
                                    "Special Notes: No Debit Card Client";
            }
            else if (AetnaHMOLayout == true && EV == true)
            {
                string eVee = "EV has been setup in BackOffice";
                string layoutGd = AetnaCarrierName;
                string noLayoutGd = "Layout Not Submitted Waiting For Reply From Carrier";
                richTextBox1.Text = "Connected Claims setup for " + subjectLine + "\n" +
                                    @"Group Folder: C:\Users\14025\Documents\File Consultants\Groups\" + subjectLine + @"\" + "\n" +
                                    "Back Office: Carrier code " + "AET_HMO" + " is setup.\n" +
                                    "File Layout Guide: " + layoutGd + "\n" +
                                    "NCC: Has been stored in this groups employer DOCS folder.\n" +
                                    "EV: " + eVee + "\n";            }
            else if (AetnaHMOLayout == true && EV == false)
            {
                string eVee = "EV hasn't been setup in BackOffice";
                string layoutGd = AetnaCarrierName;
                string noLayoutGd = "Layout Not Submitted Waiting For Reply From Carrier";
                richTextBox1.Text = "Connected Claims setup for " + subjectLine + "\n" +
                                    @"Group Folder: C:\Users\14025\Documents\File Consultants\Groups\" + subjectLine + @"\" + "\n" +
                                    "Back Office: Carrier code " + "AET_HMO" + " is setup.\n" +
                                    "File Layout Guide: " + layoutGd + "\n" +
                                    "NCC: Has been stored in this groups employer DOCS folder.\n" +
                                    "EV: " + eVee + "\n" +
                                    "Special Notes: No Debit Card Client";
            }


            else if (existingLayout == true)
            {
                string eVee = "EV has been setup in BackOffice";
                //string layoutGd = "";
                string noLayoutGd = "Existing Layout in place.";
                richTextBox1.Text = "Connected Claims setup for " + subjectLine + "\n" +
                                    @"Group Folder: C:\Users\14025\Documents\File Consultants\Groups\" + subjectLine + @"\" + "\n" +
                                    "Back Office: Carrier code " + backOfficeCode + " is setup.\n" +
                                    "File Layout Guide: " + noLayoutGd + "\n" +
                                    "NCC: Has been stored in this groups employer DOCS folder.\n" +
                                    "EV: " + eVee + "\n";
            }



            else if (payFlexLayout == true && EV == true)
            {
                string eVee = "EV has been setup in BackOffice";
                string layoutGd = "Carrier File Guide.pdf";
                richTextBox1.Text = "Connected Claims setup for " + subjectLine + "\n" +
                                    @"Group Folder: C:\Users\14025\Documents\File Consultants\Groups\" + subjectLine + @"\" + "\n" +
                                    "Back Office: Carrier code " + backOfficeCode + " is setup.\n" +
                                    "File Layout Guide: " + layoutGd + "\n" +
                                    "NCC: Has been stored in this groups employer DOCS folder.\n" +
                                    "EV: " + eVee + "\n";
            }
            else if (payFlexLayout == true && EV == false)
            {
                string eVee = "EV hasn't been setup in BackOffice";
                string layoutGd = "Carrier File Guide.pdf";
                richTextBox1.Text = "Connected Claims setup for " + subjectLine + "\n" +
                                    @"Group Folder: C:\Users\14025\Documents\File Consultants\Groups\" + subjectLine + @"\" + "\n" +
                                    "Back Office: Carrier code " + backOfficeCode + " is setup.\n" +
                                    "File Layout Guide: " + layoutGd + "\n" +
                                    "NCC: Has been stored in this groups employer DOCS folder.\n" +
                                    "EV: " + eVee + "\n";
            }
            else if (payFlexLayout == false && EV == true && layoutPrvded == true)
            {
                string eVee = "EV has been setup in BackOffice";
                //string layoutGd = "some file name";
                string noLayoutGd = "Layout Not Submitted Waiting For Reply From Carrier";
                richTextBox1.Text = "Connected Claims setup for " + subjectLine + "\n" +
                                    @"Group Folder: C:\Users\14025\Documents\File Consultants\Groups\" + subjectLine + @"\" + "\n" +
                                    "Back Office: Carrier code " + backOfficeCode + " is setup.\n" +
                                    "File Layout Guide: " + LayoutFileName + "\n" +
                                    "NCC: Has been stored in this groups employer DOCS folder.\n" +
                                    "EV: " + eVee + "\n";

            }
            else if (payFlexLayout == false && EV == false && layoutPrvded == true)
            {
                string eVee = "EV hasn't been setup in BackOffice";
                //string layoutGd = "some file name";
                string noLayoutGd = "Layout Not Submitted Waiting For Reply From Carrier";
                richTextBox1.Text = "Connected Claims setup for " + subjectLine + "\n" +
                                    @"Group Folder: C:\Users\14025\Documents\File Consultants\Groups\" + subjectLine + @"\" + "\n" +
                                    "Back Office: Carrier code " + backOfficeCode + " is setup.\n" +
                                    "File Layout Guide: " + LayoutFileName + "\n" +
                                    "NCC: Has been stored in this groups employer DOCS folder.\n" +
                                    "EV: " + eVee + "\n";
                
            }
            else if (payFlexLayout == false && EV == true && layoutPrvded == false)
            {
                string eVee = "EV has been setup in BackOffice";
                //string layoutGd = "some file name";
                string noLayoutGd = "Layout Not Submitted Waiting For Reply From Carrier";
                richTextBox1.Text = "Connected Claims setup for " + subjectLine + "\n" +
                                    @"Group Folder: C:\Users\14025\Documents\File Consultants\Groups\" + subjectLine + @"\" + "\n" +
                                    "Back Office: Carrier code " + backOfficeCode + " is setup.\n" +
                                    "File Layout Guide: " + noLayoutGd + "\n" +
                                    "NCC: Has been stored in this groups employer DOCS folder.\n" +
                                    "EV: " + eVee + "\n";

            }
            else if (payFlexLayout == false && EV == false && layoutPrvded == false)
            {
                string eVee = "EV hasn't been setup in BackOffice";
                //string layoutGd = "";
                string noLayoutGd = "Layout Not Submitted Waiting For Reply From Carrier";
                richTextBox1.Text = "Connected Claims setup for " + subjectLine + "\n" +
                                    @"Group Folder: C:\Users\14025\Documents\File Consultants\Groups\" + subjectLine + @"\" + "\n" +
                                    "Back Office: Carrier code " + backOfficeCode + " is setup.\n" +
                                    "File Layout Guide: " + noLayoutGd + "\n" +
                                    "NCC: Has been stored in this groups employer DOCS folder.\n" +
                                    "EV: " + eVee + "\n";
            }
            







        }

        private void button1_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void textBox1_CtrlA_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == Convert.ToChar(1))
            {
                TextBox txt = sender as TextBox;
                txt.SelectAll();
                e.Handled = true;
            }
        }
    }
}
