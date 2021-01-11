using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;

namespace FCHelper_v001
{
    class Class_CreateCCFolders
    {
        public string CreateCCFoldersMethod(string directoryname, string carriername)
        {

            string wholeName = directoryname + " - " + carriername;
            string groupFolderPath = @"\\phx-fs-02.payflex.com\GDrive\DataServicesGroup\Personal\Brandon\Brandon's Staging Folder\" + wholeName;
            string fullPathWithDoneFolder = @"\\phx-fs-02.payflex.com\GDrive\DataServicesGroup\Personal\Brandon\Brandon's Staging Folder\" + wholeName + "\\Done";
            string fullPathWithDocsFolder = @"\\phx-fs-02.payflex.com\GDrive\DataServicesGroup\Personal\Brandon\Brandon's Staging Folder\" + wholeName + "\\Docs";
            string fullPathWithCONNECTEDCLAIMSFile = @"\\phx-fs-02.payflex.com\GDrive\DataServicesGroup\Personal\Brandon\Brandon's Staging Folder\" + wholeName + "\\CONNECTED CLAIMS.txt";



            if (!Directory.Exists(groupFolderPath))
            {
                System.IO.Directory.CreateDirectory(groupFolderPath);
                System.IO.Directory.CreateDirectory(fullPathWithDoneFolder);
                System.IO.Directory.CreateDirectory(fullPathWithDocsFolder);
                var file = File.CreateText(fullPathWithCONNECTEDCLAIMSFile);
                file.Close();

            }
            return fullPathWithDocsFolder;
        }
    }
}
