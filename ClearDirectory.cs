using System.Windows.Forms;
using System.IO;

namespace FCHelper_v001
{
    class ClearDirectory
    {
        public void ClearDirectoryMethod(string directoryPath)
        {
            PrivateClearDirectoryMethod(directoryPath);
        }

        private void PrivateClearDirectoryMethod(string directoryPath)
        {
            string[] filePaths = Directory.GetFiles(directoryPath);
            foreach (string filePath in filePaths)
            {
                try
                {
                    File.Delete(filePath);
                }
                catch (System.IO.IOException ex)
                {
                    MessageBox.Show("Method: ClearDirectoryMethod()\r\r" + ex, "Could not delete employer directory", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                catch (System.ArgumentException ex)
                {
                    MessageBox.Show("Method: ClearDirectoryMethod()\r\r" + ex, "Could not delete employer directory", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                catch (System.NotSupportedException ex)
                {
                    MessageBox.Show("Method: ClearDirectoryMethod()\r\r" + ex, "Could not delete employer directory", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }            


            }
            
        }
    }
}
