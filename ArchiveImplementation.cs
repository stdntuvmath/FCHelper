using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Diagnostics;
using System.Windows.Forms;
using System.Diagnostics;


namespace FCHelper_v001
{
    class ArchiveImplementation
    {

        public void ArchiveImplementationMethod(string ERname, string ERID)
        {
            PrivateArchiveImplementationMethod(ERname, ERID);
        }

        private void PrivateArchiveImplementationMethod(string ername, string erid)
        {

            try
            {
                //create archive folder
                DateTime date = DateTime.Today;
                string filePathWithDate = @"C:\Users\14025\Documents\File Consultants\Brandon\Archive\" + date.ToString("yyyyMMdd_") + ername + "_" + erid + "_Completed";

                System.IO.Directory.CreateDirectory(filePathWithDate);
                
            }
            catch (Exception ex)
            {
                MessageBox.Show("Method: ArchiveImplementation\rSomething prevented the Archive folder from creating.\r\r" + ex,"Error",MessageBoxButtons.OK,MessageBoxIcon.Error);
            }


            try
            {
                //find the selected notes and move them to archive folder
                string partialName = erid;
                DirectoryInfo notesFolder = new DirectoryInfo(@"C:\Users\14025\Documents\File Consultants\Brandon\Notes\");
                FileInfo[] filesInNotes = notesFolder.GetFiles("*" + partialName + "*.*");

                foreach (FileInfo foundFile in filesInNotes)
                {
                    try
                    {
                        string fullName = foundFile.FullName;
                        //copy file to archive folder
                        string fileToCopy = @"C:\Users\14025\Documents\File Consultants\Brandon\Notes\" + ername + "_" + erid + "Notes";
                        string destinationDirectory = @"C:\Users\14025\Documents\File Consultants\Brandon\Archive\";

                        File.Copy(fileToCopy, destinationDirectory + Path.GetFileName(fileToCopy));
                        break;

                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Method: ArchiveImplementation\rSomething prevented the notes file from copying to the Archive folder\r\r" + ex, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        
                    }

                    try
                    {
                        //delete notes file in notes folder
                        string fileToCopy = @"C:\Users\14025\Documents\File Consultants\Brandon\Notes\" + ername + "_" + erid + "Notes";

                        File.Delete(fileToCopy);
                        break;
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Method: ArchiveImplementation\rSomething prevented the notes file from deleting\r\r" + ex, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);

                    }

                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Method: ArchiveImplementation\rSomething prevented the notes file from being found.\r\r" + ex, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

           

            
           // File.WriteAllText(ername, inputText);

            
        }


      
    }

}

