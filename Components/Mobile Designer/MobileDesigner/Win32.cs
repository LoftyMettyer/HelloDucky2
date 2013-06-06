using System;
using System.Runtime.InteropServices;

namespace MobileDesigner
{
    public class Win32
    {
        [DllImport("user32.dll")]
        public static extern Int32 SystemParametersInfo(uint uiAction, uint uiParam, ref bool pvParam, uint fWinIni);
        public const int SPI_SETKEYBOARDCUES = 0x100B;
    }
}
