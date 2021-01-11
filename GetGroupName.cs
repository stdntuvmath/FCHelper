using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;

namespace FCHelper_v001
{
    class GetGroupName
    {


        public static string GroupName;


        public void GetGroupNameMethod(string erid)
        {
            
            //search groups folder for ERID, if !Exist then catch



            DirectoryInfo groupsDirectory = new DirectoryInfo(@"C:\Users\14025\Documents\File Consultants\Groups");
            DirectoryInfo[] directoryNames = groupsDirectory.GetDirectories("*" + erid + "*.*");//gets all the directories with the ERID in it


            


            int counter = 0;
            foreach (DirectoryInfo name in directoryNames)
            {
                string groupName = name.Name;

                string strToCompare = "-";
                string str = groupName;

                if (str.Split().Count(r => r == strToCompare) > 1)
                {
                    //do nothing and go to the next iteration
                }
                else
                {
                    GroupName = groupName;//delivers goupName without carriers
                }
                               



            }
            


        }
    }
}
