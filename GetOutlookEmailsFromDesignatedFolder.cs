using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Outlook;
using System.Windows.Forms;


namespace FCHelper_v001
{
    class GetOutlookEmailsFromDesignatedFolder
    {


        private static string EmailSubjectLine;
        private static string EmailBody;
        private static string ERID;
        private static MAPIFolder SubFolderETLObject;
        private static string SubFolderETLString;







        public void GetOutlookEmailsFromDesignatedFolderMethod(string InboxFolderName)
        {
            PrivateGetOutlookEmailsFromDesignatedFolderMethod(InboxFolderName);
        }

        private void PrivateGetOutlookEmailsFromDesignatedFolderMethod(string InboxFolderName)
        {
            Microsoft.Office.Interop.Outlook.Application outlook = new Microsoft.Office.Interop.Outlook.Application();
            NameSpace outlookNameSpace = outlook.GetNamespace("MAPI");
            MAPIFolder rootFolder = outlookNameSpace.DefaultStore.GetRootFolder();
            MAPIFolder defaultFolder = outlookNameSpace.GetDefaultFolder(OlDefaultFolders.olFolderInbox);
            
            try
            {
                
                try
                {
                    MAPIFolder subFolderETL = defaultFolder.Folders[InboxFolderName];
                    SubFolderETLObject = subFolderETL;
                    SubFolderETLString = subFolderETL.ToString();
                }
                catch (System.Exception ex)
                {
                    MessageBox.Show("The subfolder "+ SubFolderETLObject + " is not found under root folder "+ defaultFolder, "",MessageBoxButtons.OK,MessageBoxIcon.Error);
                }




                //MessageBox.Show(rootFolder.Name);

                //Get folder other than default inbox



                for (int i = 1; i <= SubFolderETLObject.Items.Count; i++)
                {
                    Microsoft.Office.Interop.Outlook.MailItem item = SubFolderETLObject.Items[i];
                    //MessageBox.Show(i.ToString());
                    //MessageBox.Show(item.Subject);
                    //MessageBox.Show(item.Body);
                    //MessageBox.Show(item.HTMLBody);

                    EmailSubjectLine = item.Subject;                    
                    EmailBody = item.Body;

                    GetStringBetweenString getBody = new GetStringBetweenString();
                    GetStringBetweenString getSubject = new GetStringBetweenString();


                    string body = getBody.GetStringBetweenStringMethod(EmailBody, "Comment", "Ticket URL");
                    string etlTicketNumberAndErid = getSubject.GetStringBetweenStringMethod(EmailSubjectLine.ToString(), "#", "-");
                    string etlTicketNumber = getSubject.GetStringBetweenStringMethod(etlTicketNumberAndErid, "", ":");

                    string erid = getSubject.GetStringBetweenStringMethod(etlTicketNumberAndErid, " "," ");
                    //string erid = getSubject.GetStringBetweenStringMethod(eridWithSpaces, " ", " ");

                    MessageBox.Show(EmailSubjectLine);

                    MessageBox.Show(body);
                    MessageBox.Show(etlTicketNumberAndErid);
                    MessageBox.Show(etlTicketNumber);
                    //MessageBox.Show(eridWithSpaces);
                    MessageBox.Show(erid);

                    //if (EmailBody.Contains("Comment"))
                    //{
                    //    MessageBox.Show(body);

                    //    if (EmailSubjectLine.Contains("Employer:"))
                    //    {
                    //        MessageBox.Show("If Subject line contains.."+erid);
                    //        ERID = erid;
                    //    }
                    //    else
                    //    {
                    //        MessageBox.Show("If Subject line contains.." + erid);
                    //        ERID = erid;
                    //    }

                    //}
                    //else
                    //{
                    //    //do nothing
                    //}


                    //get email contacts from the implementation spreadsheet
                    //MessageBox.Show("ERID going into GetSpecificDataFromExcelDatabase: " + ERID);
                    GetSpecificDataFromExcelDatabase getData = new GetSpecificDataFromExcelDatabase();
                    getData.GetSpecificDataFromExcelDatabaseMethod(ERID);



                    //create email
                    string emailBody = String.Format("<p style = \"font-size:11pt;\">Hello Everyone,<br/><br/>" +
                                                     "Test results are below:<br/><br/></p> ");


                    CreateTestResultOutlookEmail createEmail = new CreateTestResultOutlookEmail();
                    //createEmail.CreateTestResultOutlookEmailMethod();


                }
            }
            catch (NullReferenceException ex)
            {

            }
            catch (ObjectDisposedException ex)
            {

            }
            catch (EntryPointNotFoundException ex)
            {

            }

        }
    }
}
