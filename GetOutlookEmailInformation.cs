using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Outlook;
using System.Windows.Forms;

namespace FCHelper_v001
{
    class GetOutlookEmailInformation
    {

        public static string EmailSenderName;


        private static MAPIFolder FolderToBeProcessed = null;

        public void GetOutlookEmailInformationMethod(string EmailFileNameOrSubjectLine, string FullInboxFolderName)
        {
            PrivateGetOutlookEmailInformationMethod(EmailFileNameOrSubjectLine, FullInboxFolderName);
        }

        private void PrivateGetOutlookEmailInformationMethod(string EmailFileNameOrSubjectLine, string FullInboxFolderName)
        {
            Microsoft.Office.Interop.Outlook.Application outlook = new Microsoft.Office.Interop.Outlook.Application();
            NameSpace outlookNameSpace = outlook.GetNamespace("MAPI");
            MAPIFolder rootFolder = outlookNameSpace.DefaultStore.GetRootFolder();

            MessageBox.Show(rootFolder.Name);

            //Get folder other than default inbox

            foreach (MAPIFolder folder in rootFolder.Folders)
            {
                if (folder.Name == FullInboxFolderName)
                {
                    //FolderToBeProcessed = folder;
                    MessageBox.Show(folder.Name);
                }
            }


            MailItem emailToBeFound = (MailItem)FolderToBeProcessed.Items.Find(EmailFileNameOrSubjectLine);

            string subjectLine = emailToBeFound.Subject;
            string body = emailToBeFound.Body;


            MessageBox.Show(subjectLine + "\r\r\r" + body);

        }
    }
}
