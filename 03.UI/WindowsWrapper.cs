using System;
using System.Windows.Forms;

namespace Eplan.EplAddin.HMX_MCNS
{
    public class WindowWrapper : IWin32Window
    {
        private IntPtr _hwnd;

        public WindowWrapper(IntPtr handle)
        {
            _hwnd = handle;
        }

        public IntPtr Handle
        {
            get { return _hwnd; }
        }
    }
}
