using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Outlook;

namespace FCHelper_v001
{
    class GetSelectedEmail
    {
        public IEnumerable<MailItem> GetSelectedEmailMethod()
        {
            foreach (MailItem email in new Microsoft.Office.Interop.Outlook.Application().ActiveExplorer().Selection)
            {
                yield return email;
            }
        }
    }
}
