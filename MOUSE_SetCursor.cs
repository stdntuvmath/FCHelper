using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.InteropServices;

namespace FCHelper_v001
{
    class MOUSE_SetCursor
    {
        [DllImport("user32.dll", EntryPoint = "SetCursorPos")]
        private static extern bool SetCursorPos(int x, int y);

        [DllImport("user32.dll")]
        private static extern void mouse_event(int dwFlags, int dx, int dy, int cButtons, int dwExtraInfo);

        

        public void MOUSE_SetCursorMethod(int X, int Y)
        {
            PrivateMOUSE_SetCursorMethod(X, Y);
        }

        private void PrivateMOUSE_SetCursorMethod(int X, int Y)
        {
            SetCursorPos(X, Y);

        }
    }
}
