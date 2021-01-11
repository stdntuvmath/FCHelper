using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Windows.Forms;
using System.Text.RegularExpressions;

namespace FCHelper_v001
{
    class GetClientID
    {

        public static string ClientID;

        private static string AugmentedEmployerName;
        public static string ERNameAbbreviation;


        public void GetClientIDMethod(string groupNameWithERID)
        {



            GetStringBetweenString get = new GetStringBetweenString();


            string employerNameWithoutERID = get.GetStringBetweenStringMethod(groupNameWithERID,"","-");
            string firstWordInEmployerName = get.GetStringBetweenStringMethod(employerNameWithoutERID, "", " ");
            string secondWordInEmployerName = get.GetStringBetweenStringMethod(employerNameWithoutERID, " ", " ");

            //get abreviation for three worded employer names

            //Regex initials = new Regex(@"[\w^,^.]");

            Regex initials = new Regex(@"(\b[a-zA-Z])[a-zA-Z,.]* ?");
            ERNameAbbreviation = initials.Replace(employerNameWithoutERID, "$1");//gets first letter of each word in a string and excludes comma and period



            //string ShortName = "";
            //employerNameWithoutERID.Split(' ').ToList().ForEach(i => ShortName += i[0].ToString());


            // string[] ERName = employerNameWithoutERID.Split(' ');
            //// List<char> charBuffer = new List<char>(); 

            // foreach (string wordsOfERName in ERName)
            // {
            //     char[] array = wordsOfERName.ToCharArray();

            //     ERNameAbbreviation = 

            // }

            //ERNameAbbreviation = charBuffer[0].ToString()+ charBuffer[1].ToString()+ charBuffer[2].ToString();
            //string employerNameAbbreviation = firstChars[0];
            //MessageBox.Show(ERNameAbbreviation);


            // int numberOfSpacesInEmployerName = employerNameWithoutERID.Count(x => x == ' ');//counting the number of underscores in filename
            //// MessageBox.Show(numberOfSpacesInEmployerName.ToString());

            // if (firstWordInEmployerName == "The")
            // {
            //     AugmentedEmployerName = secondWordInEmployerName;
            // }
            // else if (numberOfSpacesInEmployerName == 0)
            // {
            //     AugmentedEmployerName = firstWordInEmployerName;
            // }
            // else if (numberOfSpacesInEmployerName == 1)
            // {
            //     AugmentedEmployerName = firstWordInEmployerName + secondWordInEmployerName;

            // }
            // else if (numberOfSpacesInEmployerName > 1)
            // {
            //     AugmentedEmployerName = ERNameAbbreviation;
            // }


            //MessageBox.Show(AugmentedEmployerName);


            //AugmentedEmployerName = firstWordInEmployerName + secondWordInEmployerName;

            string ETLFolder = @"\\phx-fs-02.payflex.com\Data\PFS\ETL_Process";

            string[] directories = Directory.GetDirectories(ETLFolder);

            foreach (string folder in directories)
            {
                if (Regex.IsMatch(folder, firstWordInEmployerName+secondWordInEmployerName , RegexOptions.IgnoreCase))
                {

                    string folderNameOnly = Path.GetFileName(folder);
                   
                    string clientID = get.GetStringBetweenStringMethod(folderNameOnly, "_", "_");

                    ClientID = clientID;
                }
                else if (Regex.IsMatch(folder, firstWordInEmployerName, RegexOptions.IgnoreCase))
                {

                    string folderNameOnly = Path.GetFileName(folder);

                    string clientID = get.GetStringBetweenStringMethod(folderNameOnly, "_", "_");

                    ClientID = clientID;
                }
            }

            

        }

    }
}
