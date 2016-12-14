using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Win32;

namespace InstallerCA
{
    class RegistryAbstractor
    {
        public RegistryAbstractor()
        {

        }
        public RegistryKey OpenOrCreateHkcuKey(string subKey)
        {
            RegistryKey rkExcelXll;
            Console.WriteLine(string.Format("Opening {0} Key ...", subKey));
            if (Registry.CurrentUser.OpenSubKey(subKey) == null)
            {
                rkExcelXll = Registry.CurrentUser.CreateSubKey(subKey);
                Console.WriteLine("... key not existing, create it.");
            }
            else
            {
                rkExcelXll = Registry.CurrentUser.OpenSubKey(subKey, true);
                Console.WriteLine("... existing key successfully retrieved.");
            }
            return rkExcelXll;
        }
    }
}
