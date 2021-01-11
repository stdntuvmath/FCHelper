using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;

namespace FCHelper_v001
{
    class SaveAndWriteToRTFFile
    {
        public void SaveAndWriteToRTFFileMethod(string txt, string file)
        {
            PrivateSaveAndWriteToRTFFileMethod(txt, file);
        }

        private void PrivateSaveAndWriteToRTFFileMethod(string inputText, string inputFilePath)
        {
            
                StreamWriter stream = new StreamWriter(inputFilePath, true);
                        
                stream.Write(inputText);

                stream.Flush();

                stream.Close();


            
          



        }
    }
}
