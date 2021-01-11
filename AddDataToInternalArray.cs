using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;


namespace FCHelper_v001
{
    class AddDataToInternalArray
    {
        public void AddDataToInternalArrayMethod(string input)
        {
            string[] allFields = { };

            for (int i=0;i<=15;i++)
            {
                allFields[i] = input;
            }
        }
    }
}
