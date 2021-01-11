using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FCHelper_v001
{
    class GetterSetterObject
    {
        private object _name;//global varibable to this class



        public object DataValue

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
