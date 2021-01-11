using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FCHelper_v001
{
    class CheckTextbox1and2
    {
        public bool CheckTextbox1and2Method(string textbox1Text, string textbox2Text)
        {
            bool Boolean_T_F = false;

            PrivateCheckTextbox1and2Method( textbox1Text, textbox2Text);

            if (PrivateCheckTextbox1and2Method(textbox1Text, textbox2Text) == true)
            {
                Boolean_T_F = true;
            }

            return Boolean_T_F;
        }

        private bool PrivateCheckTextbox1and2Method(string textbox1Text, string textbox2Text)
        {
            bool boolean_T_F = false;
            if (textbox1Text == "" && textbox2Text == "")
            {
                boolean_T_F = true;
            }

            return boolean_T_F;
        }
    }
}
