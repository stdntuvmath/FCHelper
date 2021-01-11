using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace FCHelper_v001
{
    class CheckIfFormIsOpen
    {        

        public bool CheckIfFormIsOpenMethod(string programName)
        {
            bool  FormOpen = PrivateCheckIfFormIsOpenMethod(programName);

            return FormOpen;
        }

        private bool PrivateCheckIfFormIsOpenMethod(string programName)
        {
            bool formOpen = Application.OpenForms.Cast<Form>().Any(form => form.Name == programName);

            return formOpen;
        }
    }
}
