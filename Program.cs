using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Security.Principal;

namespace FCHelper_v001
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            //string folderPath = @"G:\\phx-fs-02.payflex.com\GDrive\DataServicesGroup\Personal\Brandon";
            //string domainName = @"payflex.com";
            //string username = "PA155965";
            //string password = "Yrthsa125";
            //MapDrive mapDrive = new MapDrive();
            //mapDrive.MapDriveMethod(folderPath, domainName, username, password);

            
            //map Brandons personal folder and all subfolders

            string folderPath1 = @"G:\\phx-fs-02.payflex.com\GDrive\DataServicesGroup\Personal\Brandon";
            string domainName1 = @"AETH";
            string username1 = WindowsIdentity.GetCurrent().Name;
            string password1 = "Yrthsa1222";
            MapDrive mapDrive1 = new MapDrive();
            mapDrive1.MapDriveMethod(folderPath1, domainName1, username1, password1);

            //Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new Form_Main());



        }
    }
}
