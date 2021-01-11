using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FCHelper_v001
{
    class ETLAutomatedEmails
    {
        public void ETLAutomatedEmailsMethod()
        {
            PrivateETLAutomatedEmailsMethod();
        }

        private void PrivateETLAutomatedEmailsMethod()
        {
            string folderName = "ETL";
            GetOutlookEmailsFromDesignatedFolder getEmails = new GetOutlookEmailsFromDesignatedFolder();
            getEmails.GetOutlookEmailsFromDesignatedFolderMethod(folderName);
            
            //monitor the ETL inbox folder



         

        }
    }
}
