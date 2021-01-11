using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO.Compression;
using System.Windows.Forms;

namespace FCHelper_v001
{
    class ZipFile
    {
        public void ZipFileMethod(string pathOfFileToZip, string zippedFileName)
        {
            PrivateZipFileMethod(pathOfFileToZip, zippedFileName);
        }
        private void PrivateZipFileMethod(string pathOfFileToZip, string zippedFileName)
        {

            try
            {
                System.IO.Compression.ZipFile.CreateFromDirectory(pathOfFileToZip, zippedFileName);
            }
            catch (System.IO.IOException ex)
            {
                MessageBox.Show("Method: ZipFileMethod()\r\r"+ex,"Could not zip/archive employer directory",MessageBoxButtons.OK,MessageBoxIcon.Error);
            }
            catch (System.ArgumentException ex)
            {
                MessageBox.Show("Method: ZipFileMethod()\r\r" + ex, "Could not zip/archive employer directory", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            catch (System.NotSupportedException ex)
            {
                MessageBox.Show("Method: ZipFileMethod()\r\r" + ex, "Could not zip/archive employer directory", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }




        }
    }
}
