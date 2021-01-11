using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Windows.Forms;

namespace FCHelper_v001
{
    class MoveFile
    {
        public void MoveFileMethod(string fullPathOfFileToMove, string fullDestinationPathAndFileName)
        {
            PrivateMoveFileMethod(fullPathOfFileToMove, fullDestinationPathAndFileName);
        }
        private void PrivateMoveFileMethod(string fileName, string destinationName)
        {
            try
            {
                File.Move(fileName, destinationName);
            }
            catch (System.IO.IOException ex)
            {
                MessageBox.Show("Method: MoveFile()\r\r"+ex,"Couldn't Move File",MessageBoxButtons.OK,MessageBoxIcon.Error);
            }
            catch (System.ArgumentException ex)
            {
                MessageBox.Show("Method: MoveFile()\r\r" + ex, "Couldn't Move File", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            catch (System.NotSupportedException ex)
            {
                MessageBox.Show("Method: MoveFile()\r\r" + ex, "Couldn't Move File", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }
    }
}
