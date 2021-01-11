
using System.Threading;
using System.Runtime.InteropServices;

namespace FCHelper_v001
{
    class MOUSE_LeftClick
    {

        [DllImport("user32.dll", EntryPoint = "SetCursorPos")]
        private static extern bool SetCursorPos(int x, int y);

        [DllImport("user32.dll")]
        private static extern void mouse_event(int dwFlags, int dx, int dy, int cButtons, int dwExtraInfo);

        private const int MOUSEEVENT_LEFTDOWN = 0x02;
        private const int MOUSEEVENT_LEFTUP = 0x04;

       


        public void MOUSE_LeftClickMethod(int X, int Y)
        {
            PrivateMOUSE_LeftClickMethod(X,Y);
        }

        private void PrivateMOUSE_LeftClickMethod(int X, int Y)
        {
            SetCursorPos(X,Y);

            mouse_event(MOUSEEVENT_LEFTDOWN,X,Y,0,0);
            Thread.Sleep(1000);
            mouse_event(MOUSEEVENT_LEFTUP, X, Y, 0, 0);

        }
    }
}
