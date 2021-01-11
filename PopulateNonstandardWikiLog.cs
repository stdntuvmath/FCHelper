using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Diagnostics;
using System.Windows.Forms;

namespace FCHelper_v001
{
    class PopulateNonstandardWikiLog
    {
        private static string ERID;


        public void PopulateNonstandardWikiLogMethod()
        {
            PrivatePopulateNonstandardWikiLogMethod();
        }


        private void PrivatePopulateNonstandardWikiLogMethod()
        {
            MOUSE_SetCursor setCursor = new MOUSE_SetCursor();
            MOUSE_LeftClick leftClick = new MOUSE_LeftClick();

            Process.Start(@"https://aetnao365.sharepoint.com/sites/ConsumerDataServices/CDSHeadwall/ReimbursementDataServicesWiki/NonstandardFileProcessRequests-Approved.aspx");

            Thread.Sleep(10000);

            leftClick.MOUSE_LeftClickMethod(601, 350);//click on the user


            /*The Sharepoint website has stopped the ability to use SendKeys() method. In order
             to bypass this, we will need to develope a workaround similar to the MouseClick
             Simulator program but for the Keyboard instead. This would be a very large and 
             time consuming project.*/


            //Thread.Sleep(15000);

            ////enter password

            //leftClick.MOUSE_LeftClickMethod(279, 605);//click new entry


            ////enter password into login prompt
            ////SendKeys.Send("Yrthsa1222");
            ////SendKeys.Send("{ENTER}");

            //Thread.Sleep(8000);

            //leftClick.MOUSE_LeftClickMethod(321, 399);//click ERID field
            //Thread.Sleep(2000);
            //leftClick.MOUSE_LeftClickMethod(321, 399);//click ERID field
            //leftClick.MOUSE_LeftClickMethod(321, 399);//click ERID field

            //foreach (string erid in GetNonstandardFileData.EmployerID)
            //{
            //    ERID = erid;
            //}

            //SendKeys.Send(ERID);

            //leftClick.MOUSE_LeftClickMethod(343, 493);//click Employer Name field
            //leftClick.MOUSE_LeftClickMethod(343, 493);//click Employer Name field
            //Thread.Sleep(2000);
            //SendKeys.Send(GetNonstandardFileData.EmployerName);


            //leftClick.MOUSE_LeftClickMethod(343, 493);//click Requestor field
            //leftClick.MOUSE_LeftClickMethod(343, 493);//click Requestor field
            //Thread.Sleep(2000);

            //SendKeys.Send(GetNonstandardFileData.Requester);





            //leftClick.MOUSE_LeftClickMethod(356, 669);//click approving manager field
            //leftClick.MOUSE_LeftClickMethod(356, 669);//click approving manager field
            //Thread.Sleep(2000);

            //SendKeys.Send(GetNonstandardFileData.ApprovingManager);



            //leftClick.MOUSE_LeftClickMethod(650, 213);//click on main screen
            //Thread.Sleep(2000);

            //SendKeys.Send("{PGDN}");
            //SendKeys.Send("{PGDN}");
            //SendKeys.Send("{PGDN}");
            //SendKeys.Send("{PGDN}");


            //leftClick.MOUSE_LeftClickMethod(367, 224);//click on approver field

            //Thread.Sleep(2000);


            //SendKeys.Send("Turner");

            //Thread.Sleep(2000);

            //SendKeys.Send("{ENTER}");

            //Thread.Sleep(1000);

            

            //setCursor.MOUSE_SetCursorMethod(299, 651);
            ////leftClick.MOUSE_LeftClickMethod(2264, 943);//click save button

            //Thread.Sleep(6000);


            //setCursor.MOUSE_SetCursorMethod(1254, 3);
            //leftClick.MOUSE_LeftClickMethod(1254, 3);//close screen


        }
    }


}
