using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Windows.Forms;

namespace FCHelper_v001
{
    class RenameFile
    {

        //This will rename the file.
        public void RenameFileMethod(string currentFilenameAndPathToChange, string newFileNameWithoutPath)
        {
            try
            {
                FileInfo file1 = new FileInfo(currentFilenameAndPathToChange);

                string fileExtension = Path.GetExtension(currentFilenameAndPathToChange);

                //MessageBox.Show("Old file path and filename: "+oldFilePathAndName);
                //MessageBox.Show("Trying to move file to: "+file1.Directory.FullName + "\\" + newFileNameWithoutPath+fileExtension);


                if (!newFileNameWithoutPath.Contains(fileExtension))
                {
                    if (!File.Exists(file1.Directory.FullName + "\\" + newFileNameWithoutPath + fileExtension))
                    {
                        file1.MoveTo(file1.Directory.FullName + "\\" + newFileNameWithoutPath + fileExtension);

                    }

                }
                else if (newFileNameWithoutPath.Contains(fileExtension))
                {

                    if (!File.Exists(file1.Directory.FullName + "\\" + newFileNameWithoutPath))
                    {
                        file1.MoveTo(file1.Directory.FullName + "\\" + newFileNameWithoutPath);

                    }

                }


            }
            catch (ArgumentNullException ex)
            {

            }
           }
    }
}
