using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FCHelper_v001
{
    class GetterSetterInt
    {
        private int _name;//global varibable to this class



        public int DataValue

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
