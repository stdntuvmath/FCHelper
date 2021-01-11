using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Outlook;
using System.Windows.Forms;
using System.Security.Principal;
using System.IO;

namespace FCHelper_v001
{
    class CreateTestResultOutlookEmail
    {

        private static string StagineFolder = @"\\phx-fs-02.payflex.com\GDrive\DataServicesGroup\Personal\Brandon\Brandon's Staging Folder\";

        public void CreateTestResultOutlookEmailMethod(string emailTo1, string emailTo2, string emailTo3, string emailTo4,
                                             string emailCC1, string emailCC2, string emailCC3, string emailCC4, 
                                             string emailSubject, string emailBody)
        {

            Microsoft.Office.Interop.Outlook.Application outlookApp = new Microsoft.Office.Interop.Outlook.Application();
            MailItem email = outlookApp.CreateItem(OlItemType.olMailItem);

            //SetMailFormat_2002_2003_2007_2010(email);//This doesn't do anything
            /*Formatting will follow through this method on the GHost side but Aetna
             side changes the font back to 10pt somehow. I tried looking in Outlooks
             settings but did not see anything.*/

            email.To = emailTo1 + ";" + emailTo2 + ";" + emailTo3 + ";" + emailTo4;
            email.CC = emailCC1 + ";" + emailCC2 + ";" + emailCC3 + ";" + emailCC4;


            email.BCC = "DSG-Responses@AETNA.com";

            email.Subject = emailSubject;

            //string s = email.Body + email.HTMLBody;
            email.BodyFormat = OlBodyFormat.olFormatHTML;
            email.Display();
            //email.Body = emailBody;
            email.HTMLBody = emailBody + email.HTMLBody;//addes the default signature
            //email.Body = emailBody;
            //email.HTMLBody = s;


            

            //save email file in Staging folder (workaround)
            //try
            //{
            //    email.SaveAs(StagineFolder, OlSaveAsType.olMSG);

            //}
            //catch (System.Exception ex)
            //{
            //    MessageBox.Show(ex.ToString());
            //}

        }

        //for keeping format - this did not work
        //private void SetMailFormat_2002_2003_2007_2010(object mail)
        //{
        //    System.Int32 mailFormat;
        //    mailFormat = Convert.ToInt32(mail.GetType().InvokeMember("BodyFormat",
        //        System.Reflection.BindingFlags.GetProperty, null, mail, null));
        //    //OlBodyFormat.olFormatUnspecified = 0 
        //    //OlBodyFormat.olFormatPlain = 1 
        //    //OlBodyFormat.olFormatHTML = 2 
        //    //OlBodyFormat.olFormatRichText = 3 
        //    if (mailFormat == 1) mailFormat = 2;
        //    mail.GetType().InvokeMember("BodyFormat",
        //        System.Reflection.BindingFlags.SetProperty, null, mail,
        //        new object[1] { mailFormat });
        //}

    }
}
