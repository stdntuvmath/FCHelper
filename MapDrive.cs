using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace FCHelper_v001
{
    class MapDrive
    {
        public void MapDriveMethod(string folderPathToAccess, string domainNameToAccess, string username, string password /*string pFlag*/)
        {
           

                //string folderPath = @"G:\\phx-fs-02.payflex.com\GDrive\DataServicesGroup\Personal\Brandon";
                //string domainName = @"payflex.com";
                //string username = "PA155965";
                //string password = "Yrthsa12$";
                string pFlag = "NO";

                System.Diagnostics.Process process = new System.Diagnostics.Process();
                System.Diagnostics.ProcessStartInfo psi = new System.Diagnostics.ProcessStartInfo();

                //entering command to map a drive on the command prompt
                //@"/C net use "driveName:folderPath /USER: DomainName\username password/PERSISTENT: yes/no "
                psi.FileName = "cmd.exe";
                psi.Arguments = @"/C net use " + folderPathToAccess + "/USER:" + domainNameToAccess + @"\" + username + @" " + password + @"/PERSISTENT:" + pFlag;


            try
            {
                psi.WindowStyle = System.Diagnostics.ProcessWindowStyle.Normal;
                process.StartInfo = psi;
            }
            catch (System.InvalidOperationException ex)
            {
                MessageBox.Show("Method: MapDrive()\r\rSomething prevented the mapping of the requested drive.\r\r" + ex);
            }
            catch (System.ArgumentException ex)
            {
                MessageBox.Show("Method: MapDrive()\r\rSomething prevented the mapping of the requested drive.\r\r" + ex);
            }

            try
            {
                process.Start();
            }
            catch (System.InvalidOperationException ex)
            {
                MessageBox.Show("Method: MapDrive()\r\rSomething prevented the mapping of the requested drive.\r\r"+ex);
            }
            catch (System.ComponentModel.Win32Exception ex)
            {
                MessageBox.Show("Method: MapDrive()\r\rSomething prevented the mapping of the requested drive.\r\r" + ex);
            }
           

        }
    }
}
