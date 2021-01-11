using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Windows.Forms;

namespace FCHelper_v001
{
    class GetErrorReportFileName
    {

        public static List<string> ErrorReportGroupFolderLocations = null;
        public static List<string> ErrorReportFileNamesOnly = null;
        public static List<string> ErrorReportTempFolderLocations = null;


        public void GetErrorReportFileNameMethod(string erid, string processedDateAndTimeStamp)
        {
            //find the main group folder name
            GetGroupName gName = new GetGroupName();
            gName.GetGroupNameMethod(erid);
            string groupName = GetGroupName.GroupName;

            string errorDirectory = @"C:\Users\14025\Documents\File Consultants\Groups\" + groupName + @"\Import\Errors";

            List<string> groupFolderLocation = new List<string>();
            List<string> fileNameOnly = new List<string>();
            List<string> tempFolderLocation = new List<string>();



            foreach( string file in Directory.GetFiles(errorDirectory, "*" + processedDateAndTimeStamp + "*", SearchOption.AllDirectories))
            {
                groupFolderLocation.Add(file);

                string fNameOnly = Path.GetFileName(file);

                fileNameOnly.Add(fNameOnly);

                string destinationFile = @"C:\Users\14025\Documents\File Consultants\Brandon\Temp\" + fNameOnly;

                tempFolderLocation.Add(destinationFile);
            }



            //foreach (string dir in Directory.GetDirectories(errorDirectory))
            //{
            //    foreach ( string f in Directory.GetFiles(dir))
            //    {
            //        if (f.Contains(processedDateAndTimeStamp))
            //        {

            //            groupFolderLocation.Add(f);

            //            string fNameOnly = Path.GetFileName(f);

            //            fileNameOnly.Add(fNameOnly);

            //            string destinationFile = @"C:\Users\14025\Documents\File Consultants\Brandon\Temp\" + fNameOnly;

            //            tempFolderLocation.Add(destinationFile);
            //        }
            //    }

            //}


            ErrorReportGroupFolderLocations = groupFolderLocation;

            ErrorReportFileNamesOnly = fileNameOnly;

            ErrorReportTempFolderLocations = tempFolderLocation;


        }
    }
}
