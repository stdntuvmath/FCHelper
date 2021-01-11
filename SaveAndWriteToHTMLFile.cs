using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;

namespace FCHelper_v001
{
    class SaveAndWriteToHTMLFile
    {
        public void SaveAndWriteToHTMLFileMethod(string txt, string filePath)
        {
            PrivateSaveAndWriteToHTMLFileMethod(txt, filePath);
        }

        private void PrivateSaveAndWriteToHTMLFileMethod(string inputText, string inputFilePath)
        {
            //StreamWriter file = new StreamWriter(inputFilePath);
            //file.Write(inputText);
            //file.Close();

            if (!File.Exists(inputFilePath))

            {

                StreamWriter stream = new StreamWriter(inputFilePath, true, System.Text.Encoding.UTF8);

                stream.Write(@"<html>" + stream.NewLine + @"<body>" + stream.NewLine);

                stream.Write(@"<p>"+ inputText + "</p>"+ stream.NewLine + "</body>" + stream.NewLine + "</html>");

                stream.Flush();

                stream.Close();

            }
            else
            {

            }


        }
    }
}
