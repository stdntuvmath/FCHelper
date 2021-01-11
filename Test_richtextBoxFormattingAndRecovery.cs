using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using Microsoft.Office.Interop.Word;

namespace FCHelper_v001
{
    public partial class Test_richtextBoxFormattingAndRecovery : Form
    {
        public Test_richtextBoxFormattingAndRecovery()
        {
            InitializeComponent();
        }

        private void highlightSelectedTextToolStripMenuItem_Click(object sender, EventArgs e)
        {
            richTextBox1.SelectionBackColor = System.Drawing.Color.Yellow;
        }

        private void button1_Click(object sender, EventArgs e)//save
        {
            string filePath = @"C:\Users\PA155965\source\repos\FCHelper v001\FCHelper v001\bin\Debug\" + textBox1.Text + "_" + textBox2.Text + "_Notes.rtf";

            string textToSave = richTextBox1.Rtf;



            SaveAndWriteToRTFFile save = new SaveAndWriteToRTFFile();
            save.SaveAndWriteToRTFFileMethod(textToSave, filePath);

            richTextBox1.SaveFile(filePath, RichTextBoxStreamType.RichText);
        }

        private void button2_Click(object sender, EventArgs e)//load
        {
            string partialName = textBox2.Text;
            DirectoryInfo notesFolder = new DirectoryInfo(@"C:\Users\PA155965\source\repos\FCHelper v001\FCHelper v001\bin\Debug\");
            FileInfo[] filesInNotes = notesFolder.GetFiles("*" + partialName + "*.*");

            foreach (FileInfo foundFile in filesInNotes)
            {
               // richTextBox1.LoadFile(foundFile.ToString(),RichTextBoxStreamType.RichText);


               

                try
                {
                    using (var rtf = new RichTextBox())
                    {
                        rtf.Rtf = File.ReadAllText(foundFile.ToString());
                        //return rtf.Text;
                        richTextBox1.Rtf = rtf.Rtf;
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Method: Test_richtexBoxFormattingAndRecovery.button2_Click\rSomething prevented the file from loading to the RickTextBox.\r\r" + ex, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);

                }
            }
        }
    }
}
