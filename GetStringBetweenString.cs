using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FCHelper_v001
{
    class GetStringBetweenString
    {
        public string GetStringBetweenStringMethod(string givenString, string partOfStringYouWantBegin, string partOfStringYouWantEnd)
        {
           
            int Start, End;
            if (givenString.Contains(partOfStringYouWantBegin) && givenString.Contains(partOfStringYouWantEnd))
            {
                Start = givenString.IndexOf(partOfStringYouWantBegin, 0) + partOfStringYouWantBegin.Length;
                End = givenString.IndexOf(partOfStringYouWantEnd, Start);
                return givenString.Substring(Start, End - Start);
            }
            else
            {
                return "";
            }
            
        }
    }
}
