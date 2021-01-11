using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;

namespace FCHelper_v001
{
    class SaveAndWriteToTextFile
    {
        public void SaveAndWriteToTextFileMethod(string txt, string file)
        {
            PrivateSaveAndWriteToTextFileMethod(txt, file);
        }

        private void PrivateSaveAndWriteToTextFileMethod(string inputText, string inputFilePath)
        {
            //StreamWriter file = new StreamWriter(inputFilePath);
            //file.Write(inputText);
            //file.Close();

            File.WriteAllText(inputFilePath, inputText);
            

        }
    }
}
