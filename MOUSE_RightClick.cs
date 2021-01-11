using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.InteropServices;

namespace FCHelper_v001
{
    class MOUSE_RightClick
    {

        [DllImport("user32.dll", EntryPoint = "SetCursorPos")]
        private static extern bool SetCursorPos(int x, int y);

        [DllImport("user32.dll")]
        private static extern void mouse_event(int dwFlags, int dx, int dy, int cButtons, int dwExtraInfo);

        

        private const int MOUSEEVENT_RIGHTDOWN = 0x08;
        private const int MOUSEEVENT_RIGHTUP = 0x10;



        public void MOUSE_RightClickMethod(int X, int Y)
        {
            PrivateMOUSE_RightClickMethod(X, Y);
        }

        private void PrivateMOUSE_RightClickMethod(int X, int Y)
        {
            SetCursorPos(X, Y);

            mouse_event(MOUSEEVENT_RIGHTDOWN, X, Y, 0, 0);
            mouse_event(MOUSEEVENT_RIGHTUP, X, Y, 0, 0);
        }
    }
}
