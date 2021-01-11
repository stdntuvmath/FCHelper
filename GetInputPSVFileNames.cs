using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Windows.Forms;

namespace FCHelper_v001
{
    class GetInputPSVFileNames
    {

        public static List<string> PSVFolderLocations = null;
        public static List<string> InputPSVFileNamesOnly = null;
        public static List<string> InputPSVTempFolderLocations = null;
        public static List<string> OutputPSVFileNamesOnly = null;

        public void GetInputPSVFileNamesMethod(string erid, string processedDateAndTimeStamp)
        {
            //find the main group folder name
            GetGroupName gName = new GetGroupName();
            gName.GetGroupNameMethod(erid);
            string groupName = GetGroupName.GroupName;



            string exportDirectory = @"C:\Users\14025\Documents\File Consultants\Groups\" + groupName + @"\Export";



            List<string> groupFolderLocation = new List<string>();
            List<string> inputfileNameOnly = new List<string>();
            List<string> outputfileNameOnly = new List<string>();
            List<string> tempFolderLocation = new List<string>();


            
            foreach (string file in Directory.GetFiles(exportDirectory, "*" + processedDateAndTimeStamp + "*", SearchOption.AllDirectories))
            {



                if (file.Contains(processedDateAndTimeStamp))
                {
                    groupFolderLocation.Add(file);

                    string fNameOnly = Path.GetFileName(file);




                    int count = fNameOnly.Count(x => x == '_');//counting the number of underscores

                    int indexOfFirstDash = fNameOnly.IndexOf('-');
                    int indexOfFirstUnderscore = fNameOnly.IndexOf('_');


                    if (count >= 5 || indexOfFirstDash == 14)
                    {
                        outputfileNameOnly.Add(fNameOnly);
                        //MessageBox.Show("This file contains 4 undescores and is an input file. " + fNameOnly);

                    }
                    else if (count < 5 || indexOfFirstDash > 20)
                    {
                        inputfileNameOnly.Add(fNameOnly);
                        //MessageBox.Show("This file contains 4 undescores and is an input file. " + fNameOnly);

                    }
                    
                    //else if (fNameOnly.Contains("-") && !fNameOnly.Contains("TEST"))
                    //{
                    //    outputfileNameOnly.Add(fNameOnly);
                    //    //MessageBox.Show("This file contains - and is an output file. "+fNameOnly);

                    //}
                    //else if (fNameOnly.Contains("-") && !fNameOnly.Contains("TEST"))
                    //{
                    //    outputfileNameOnly.Add(fNameOnly);
                    //    //MessageBox.Show("This file contains - and is an output file. "+fNameOnly);

                    //}
                    


                    //MessageBox.Show("This file contains 5 underscores and is an output file. " + fNameOnly);
                    


                    string destinationFile = @"C:\Users\14025\Documents\File Consultants\Brandon\Temp\" + fNameOnly;

                    tempFolderLocation.Add(destinationFile);
                }


                PSVFolderLocations = groupFolderLocation;

                InputPSVFileNamesOnly = inputfileNameOnly;

                OutputPSVFileNamesOnly = outputfileNameOnly;

                InputPSVTempFolderLocations = tempFolderLocation;

            }
        }

    }
}
