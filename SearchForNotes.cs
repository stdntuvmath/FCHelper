using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;

namespace FCHelper_v001
{
    class SearchForNotes
    {
        public string SearchForNotesMethod(string erid)
        {

            
                DirectoryInfo dirPath = new DirectoryInfo(@"C:\Users\14025\Documents\File Consultants\Brandon\Notes");

                FileInfo[] filesInDir = dirPath.GetFiles("*" + erid + "*.*");

         

                var fileName = filesInDir.First().ToString();//get the first filename with ERID
                return fileName;                    
            
        }
    }
}
