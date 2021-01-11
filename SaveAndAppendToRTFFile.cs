using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;

namespace FCHelper_v001
{
    class SaveAndAppendToRTFFile
    {
        public void SaveAndAppendToRTFFileMethod(string txt, string file)
        {
            PrivateSaveAndAppendToRTFFileMethod(txt, file);
        }

        private void PrivateSaveAndAppendToRTFFileMethod(string inputText, string inputFilePath)
        {

            StreamWriter stream = new StreamWriter(inputFilePath, true);

            stream.Write(inputText);

           

            stream.Flush();

            stream.Close();

        }
    }
}
