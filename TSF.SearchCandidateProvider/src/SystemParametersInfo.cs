using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace TSF.SearchCandidateProviderInternal
{
    internal static class SystemParametersInfo
    {
        private const uint SPI_GETTHREADLOCALINPUTSETTINGS = 0x104E;
        private const uint SPI_SETTHREADLOCALINPUTSETTINGS = 0x104F;

        [return: MarshalAs(UnmanagedType.Bool)]
        [DllImport("user32.dll", EntryPoint = "SystemParametersInfoW")]
        private extern static bool SystemParametersInfoW(uint action, uint param, [MarshalAs(UnmanagedType.Bool)] out bool value, uint notused);

        [return: MarshalAs(UnmanagedType.Bool)]
        [DllImport("user32.dll", EntryPoint = "SystemParametersInfoW")]
        private extern static bool SystemParametersInfoW(uint action, uint param, [MarshalAs(UnmanagedType.SysInt)] IntPtr value, uint notused);

        public static bool GetThreadLocalInputSettings(bool defaultValue)
        {
            bool result = defaultValue;
            if (!SystemParametersInfoW(SPI_GETTHREADLOCALINPUTSETTINGS, 0, out result, 0))
            {
                return defaultValue;
            }
            return result;
        }
        public static void SetThreadLocalInputSettings(bool value)
        {
            var v = IntPtr.Zero;
            if (value)
            {
                v = IntPtr.Add(v, 1);
            }
            SystemParametersInfoW(SPI_SETTHREADLOCALINPUTSETTINGS, 0, v, 0);
        }
    }
}
