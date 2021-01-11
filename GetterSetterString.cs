using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FCHelper_v001
{
    class GetterSetterString
    {
        private string _name;//global varibable to this class



        public string DataValue

        {
            get
            {
                return _name;
            }

            set
            {
                _name = value;
            }


        }
    }
}
