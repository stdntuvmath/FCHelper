using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Diagnostics;
using System.Windows.Forms;
using System.Threading;
using System.Windows;
using System.Runtime.InteropServices;

namespace FCHelper_v001
{
    class PushButtonsOnPFM_ContactChange
    {

        //force PFM into focus

        
        //call SetForegroundWindow method
        [DllImport("user32.dll")]
        public static extern IntPtr SetForegroundWindow(IntPtr hWnd);

        //turn on PFM if it isn't on already

        public void ActiveatePFMMethod()
        {
            string pfmActiveName = @"C:\Users\PA155965\Desktop\PayFlex File Manager.appref-ms";

           
            Process getPFM = Process.Start(pfmActiveName);
            Thread.Sleep(200);           

            SendKeys.Send("{TAB}{TAB}");
            Thread.Sleep(500);
            SendKeys.Send("{ENTER}");
            SendKeys.Send("{ENTER}");
            SendKeys.Send("{ENTER}");

            //getPFM.WaitForExit(5000);

           // SendKeys.Send("%");

           // Thread.Sleep(5000);

            PushButtonsOnPFM_ContactChangeMethod();
        }



        //select menu option on PFM
        private void PushButtonsOnPFM_ContactChangeMethod()
        {
            
            //SetForegroundWindow(getPFM[0].Handle);
        
            
            SendKeys.SendWait("%");
            //SendKeys.SendWait("%");
            //SendKeys.SendWait("%");
            //SendKeys.SendWait("%");
            //SendKeys.SendWait("%");
            //SendKeys.Send("{RIGHT}");


        }

    }
}
