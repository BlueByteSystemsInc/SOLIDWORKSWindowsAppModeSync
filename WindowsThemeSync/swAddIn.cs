using Microsoft.Win32;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using SolidWorks.Interop.swpublished;
using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace WindowsThemeSync
{
    [Guid("36EB446E-E8D2-4F2A-AE07-B45D8DD92694"), ComVisible(true)]
    public class WindowsThemeSyncAddIn : ISwAddin
    {
        public SldWorks SOLIDWORKS { get; private set; }
        public int SessionCookie { get; private set; }
        public static object AddInName { get; private set; } = "Windows 10 App Mode Sync - Blue Byte Systems Inc.";
        public static object AddInDescription { get; private set; } = "Syncs your SOLIDWORKS theme with Windows 10 default app mode.";

        public bool ConnectToSW(object ThisSW, int Cookie)
        {
            try
            {
                SOLIDWORKS = ThisSW as SldWorks;
                SessionCookie = Cookie;
                SOLIDWORKS.SetAddinCallbackInfo(0, this, SessionCookie);

                RefreshTheme();
                WindowsThemeManager.WindowsThemeChanged += WindowsThemeManager_OfficeThemeChanged;
                WindowsThemeManager.StartWatchingForThemeChanges();

                return true;
            }
            catch (Exception e)
            {
                MessageBox.Show($"Message = {e.Message} Stacktrace: {e.StackTrace}");

                // Todo: log ex
                return false;
            }
        }

        private void WindowsThemeManager_OfficeThemeChanged(object sender, Theme_e e)
        {

            var currentTheme = (swInterfaceBrightnessTheme_e)SOLIDWORKS.GetUserPreferenceIntegerValue((int)swUserPreferenceIntegerValue_e.swSystemColorsBackground);

            switch (e)
            {
                case Theme_e.Light:
                    if (currentTheme != swInterfaceBrightnessTheme_e.swInterfaceBrightnessTheme_Light)
                    SOLIDWORKS.SetUserPreferenceIntegerValue((int)swUserPreferenceIntegerValue_e.swSystemColorsBackground, (int)swInterfaceBrightnessTheme_e.swInterfaceBrightnessTheme_Light);
                    break;
                case Theme_e.Dark:

                    if (currentTheme != swInterfaceBrightnessTheme_e.swInterfaceBrightnessTheme_Dark)
                        SOLIDWORKS.SetUserPreferenceIntegerValue((int)swUserPreferenceIntegerValue_e.swSystemColorsBackground, (int)swInterfaceBrightnessTheme_e.swInterfaceBrightnessTheme_Dark);
                    break;
                default:
                    break;
            }
        }


        public void RefreshTheme()
        {
            var theme = WindowsThemeManager.GetWindowsTheme();
            var currentTheme = (swInterfaceBrightnessTheme_e)SOLIDWORKS.GetUserPreferenceIntegerValue((int)swUserPreferenceIntegerValue_e.swSystemColorsBackground);

            switch (theme)
            {
                case Theme_e.Light:
                    if (currentTheme != swInterfaceBrightnessTheme_e.swInterfaceBrightnessTheme_Light)
                        SOLIDWORKS.SetUserPreferenceIntegerValue((int)swUserPreferenceIntegerValue_e.swSystemColorsBackground, (int)swInterfaceBrightnessTheme_e.swInterfaceBrightnessTheme_Light);
                    break;
                case Theme_e.Dark:

                    if (currentTheme != swInterfaceBrightnessTheme_e.swInterfaceBrightnessTheme_Dark)
                        SOLIDWORKS.SetUserPreferenceIntegerValue((int)swUserPreferenceIntegerValue_e.swSystemColorsBackground, (int)swInterfaceBrightnessTheme_e.swInterfaceBrightnessTheme_Dark);
                    break;
                default:
                    break;
            }
        }
        public bool DisconnectFromSW()
        {
            try
            {
                WindowsThemeManager.StopWatchingForThemeChanges();
                return true;
            }
            catch (Exception e)
            {
                //todo: log ex
                return false;
            }
        }


        [ComRegisterFunction]
        private static void RegisterAssembly(Type t)
        {
            try
            {
                string KeyPath = string.Format(@"SOFTWARE\SolidWorks\AddIns\{0:b}", t.GUID);
                RegistryKey rk = Registry.LocalMachine.CreateSubKey(KeyPath);
                rk.SetValue("Title", AddInName); // Title
                rk.SetValue("Description", AddInDescription); // Description
                rk.SetValue(null, 1); // startup parameter (loads addin automatically upon SOLIDWORKS startup)
            }
            catch (Exception e)
            {
                //todo: msgbox ex
                throw e;
            }
        }

        [ComUnregisterFunction]
        private static void UnregisterAssembly(Type t)
        {
            try
            {
                bool Exist = false;
                string KeyPath = string.Format(@"SOFTWARE\SolidWorks\AddIns\{0:b}", t.GUID);
                using (RegistryKey Key = Registry.LocalMachine.OpenSubKey(KeyPath))
                {
                    if (Key != null)
                        Exist = true;
                    else
                        Exist = false;
                }
                if (Exist)
                    Registry.LocalMachine.DeleteSubKeyTree(KeyPath);
            }
            catch (Exception e)
            {
                //todo: msgbox ex
                throw e;
            }
        }
    }
}
