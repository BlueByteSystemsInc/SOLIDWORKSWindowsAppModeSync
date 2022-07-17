using Microsoft.Win32;
using System;
using System.Management;
using System.Security.Principal;

namespace WindowsThemeSync
{

    public enum Theme_e
    {
        Undefined,
        Light,
        Dark
    }

    public class WindowsThemeManager
    {
        #region Public Events

        public static event EventHandler<Theme_e> WindowsThemeChanged;

        #endregion

        #region Private Fields

        private static ManagementEventWatcher watcher;

        #endregion

        #region Private Properties

      

        #endregion

        #region Public Properties

        public static bool IsWatching { get; private set; } = false;

        #endregion

        #region Private Constructors

        private WindowsThemeManager()
        {
        }

        #endregion

        #region Private Methods
        private const string RegistryKeyPath = @"Software\Microsoft\Windows\CurrentVersion\Themes\Personalize";

        private const string RegistryValueName = "AppsUseLightTheme";

        public static Theme_e GetWindowsTheme()
        {
            try
            {
                var key = Registry.CurrentUser.OpenSubKey(RegistryKeyPath);
                var registryValueObject = key?.GetValue(RegistryValueName);
                if (registryValueObject == null)
                {
                    return Theme_e.Light;
                }

                var registryValue = (int)registryValueObject;

                key.Close();

                return registryValue > 0 ? Theme_e.Light : Theme_e.Dark;
            }
            catch (Exception)
            {

                return Theme_e.Undefined;
            }
          


        }
        private static void Watcher_EventArrived(object sender, EventArrivedEventArgs e)
        {
            try
            {
                var ret = GetWindowsTheme();
                WindowsThemeChanged?.Invoke(null, ret);
            }
            catch (System.Exception)
            {
                // needs to add logger here
            }
        }

        #endregion

        #region Public Methods

         
        public static void StartWatchingForThemeChanges()
        {
            try
            {
                if (IsWatching == true)
                    throw new Exception("Object already watching registry entry");

                var currentUser = WindowsIdentity.GetCurrent();
                WqlEventQuery query = new WqlEventQuery(
                "SELECT * FROM RegistryValueChangeEvent WHERE " +
                "Hive = 'HKEY_USERS' " +
                $@"AND KeyPath='{currentUser.User.Value}\\Software\\Microsoft\\Windows\\CurrentVersion\\Themes\\Personalize' AND ValueName='{RegistryValueName}'");

                watcher = new ManagementEventWatcher(query);
                watcher.EventArrived -= Watcher_EventArrived;
                watcher.EventArrived += Watcher_EventArrived;

                watcher.Start();

                IsWatching = true;
            }
            catch (Exception)
            {

                
            }
           
        }

        public static void StopWatchingForThemeChanges()
        {
            if (watcher != null)
            {
                watcher.Stop();
                watcher.Dispose();
                watcher = null;
            }
            IsWatching = false;
        }

        #endregion
    }
}
