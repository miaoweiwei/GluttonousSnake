using Microsoft.Deployment.WindowsInstaller;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Globalization;
using System.Windows.Forms;

namespace InstallerCA
{
    public class CustomActions
    {
        private const string BaseAddInKey = @"Software\Microsoft\Office\";
        private static Session pulicSession;
        private static string regasmPath;

        private static int nowVersion = 2350;//2.3.5.0之前用winform安装程序
        #region Methods

        #region CaRegisterAddIn
        [CustomAction]
        public static ActionResult CaRegisterAddIn(Session session)
        {
            //MessageBox.Show("加载注册表");
            pulicSession = session;
            bool foundOffice = false;

            session.Log("开始注册加载XLL环节（CaRegisterAddIn）......");

            try
            {
                session.Log("整理属性数据,并结构化......");
                Parameters parameters = Parameters.ExtractFromSession(session);

                var registryAdapator = new RegistryAbstractor();

                foreach (string officeVersionKey in parameters.SupportedOfficeVersion)
                {
                    double version = double.Parse(officeVersionKey, NumberStyles.Any, CultureInfo.InvariantCulture);

                    session.Log("检索注册表：{0}", BaseAddInKey + officeVersionKey);

                    string excelBaseKey = BaseAddInKey + officeVersionKey + @"\Excel";

                    if (IsOfficeExcelInstalled(excelBaseKey))
                    {
                        if (!foundOffice) foundOffice = true;
                        string excelOptionKey = excelBaseKey + @"\Options";
                        using (RegistryKey rkExcelXll = registryAdapator.OpenOrCreateHkcuKey(excelOptionKey))
                        {
                            string xllToRegister = GetAddInName(parameters.Xll32Name, parameters.Xll64Name, officeVersionKey, version);
                            session.Log("GetAddInName获取值为：{0}", xllToRegister);

                            if (xllToRegister == parameters.Xll32Name)
                            {
                                session["OFFICEBITNESS"] = "x86";
                            }
                            else if (xllToRegister == parameters.Xll64Name)
                            {
                                session["OFFICEBITNESS"] = "x64";
                            }
                            string fullPathToXll = Path.Combine(parameters.InstallDirectory, xllToRegister);

                            session.Log("在注册表HKCU中成功检索到: " + excelOptionKey);

                            string[] valueNames = rkExcelXll.GetValueNames();
                            bool isOpen = false;
                            int maxOpen = -1;
                            foreach (string valueName in valueNames)
                            {
                                session.Log(string.Format("检索 value {0}", valueName));

                                if (valueName.StartsWith("OPEN"))
                                {
                                    int openVersion = int.TryParse(valueName.Substring(4), out openVersion) ? openVersion : 0;
                                    int newOpen = valueName == "OPEN" ? 0 : openVersion;
                                    if (newOpen > maxOpen)
                                    {
                                        maxOpen = newOpen;
                                    }

                                    if (rkExcelXll.GetValue(valueName).ToString().Contains(xllToRegister))
                                    {
                                        isOpen = true;
                                        session.Log("已经发现 OPEN key " + excelOptionKey);
                                    }
                                }
                            }
                            if (!isOpen)
                            {
                                string value = "/R \"" + fullPathToXll + "\"";
                                string keyToUse;
                                if (maxOpen == -1)
                                {
                                    keyToUse = "OPEN";
                                }
                                else
                                {
                                    keyToUse = "OPEN" + (maxOpen + 1).ToString(CultureInfo.InvariantCulture);

                                }
                                rkExcelXll.SetValue(keyToUse, value);
                                session.Log("为 {0} 赋值 {1}", keyToUse, value);
                            }
                        }
                    }
                    else
                    {
                        Console.WriteLine("在注册表HKLM中没有检索到: {0}. 这个版本的Office可能没有安装", excelBaseKey);
                    }
                }
                session.Log("结束注册加载XLL环节（CaRegisterAddIn）......");
            }
            catch (System.Security.SecurityException ex)
            {
                session.Log("异常：CaRegisterAddIn SecurityException" + ex.Message);
                foundOffice = false;
            }
            catch (System.UnauthorizedAccessException ex)
            {
                session.Log("异常：CaRegisterAddIn UnauthorizedAccessException" + ex.Message);
                foundOffice = false;
            }
            catch (Exception ex)
            {
                session.Log("异常：CaRegisterAddIn Exception" + ex.Message);
                foundOffice = false;
            }
            if (!foundOffice)
            {
                MessageBox.Show("安装过程中出现错误，请关闭杀毒软件后再次尝试安装", "SumscopeAddIn");
            }
            return foundOffice ? ActionResult.Success : ActionResult.Failure;
        }
        #endregion

        #region CaUnRegisterAddIn
        [CustomAction]
        public static ActionResult CaUnRegisterAddIn(Session session)
        {
            //MessageBox.Show("卸载注册表");
            pulicSession = session;

            bool foundOffice = false;

            try
            {
                session.Log("开始注销卸载XLL环节（CaUnRegisterAddIn）......");

                session.Log("整理属性数据,并结构化......");
                Parameters parameters = Parameters.ExtractFromSession(session);

                if (parameters.SupportedOfficeVersion.Count > 0)
                {
                    foreach (string officeVersionKey in parameters.SupportedOfficeVersion)
                    {
                        string officeKey = BaseAddInKey + officeVersionKey;
                        session.Log("试图在HKCU中打开{0}", officeKey);

                        if (Registry.CurrentUser.OpenSubKey(officeKey, false) != null)
                        {
                            foundOffice = true;

                            string keyName = BaseAddInKey + officeVersionKey + @"\Excel\Options";
                            session.Log("试图在HKCU中打开{0}", keyName);

                            using (RegistryKey rkAddInKey = Registry.CurrentUser.OpenSubKey(keyName, true))
                            {
                                if (rkAddInKey != null)
                                {
                                    session.Log("存在：{0}", keyName);
                                    string[] valueNames = rkAddInKey.GetValueNames();
                                    foreach (string valueName in valueNames)
                                    {
                                        //unregister both 32 and 64 xll
                                        if (valueName.StartsWith("OPEN") && (rkAddInKey.GetValue(valueName).ToString().Contains(parameters.Xll64Name) || rkAddInKey.GetValue(valueName).ToString().Contains(parameters.Xll32Name)))
                                        {
                                            Console.WriteLine("删除： {0}", valueName);
                                            rkAddInKey.DeleteValue(valueName);
                                        }
                                    }
                                }
                                else
                                {
                                    session.Log("不存在：{0}", keyName);
                                }
                            }
                        }
                        else
                        {
                            session.Log("不存在：{0}", officeKey);
                        }
                    }
                }
                session.Log("结束注销卸载XLL环节（CaUnRegisterAddIn）......");
            }
            catch (Exception ex)
            {
                session.Log(ex.Message);
            }
            if (!foundOffice)
            {
                MessageBox.Show("安装过程中出现错误，请关闭杀毒软件后再次尝试安装", "SumscopeAddIn");
            }
            return foundOffice ? ActionResult.Success : ActionResult.Failure;
        }
        #endregion

        #region ClosePrompt
        [CustomAction]
        public static ActionResult ClosePrompt(Session session)
        {
            //MessageBox.Show("检查关闭");
            session.Log("Begin PromptToCloseApplications");
            try
            {
                var productName = session["ProductName"];
                var processes = session["PromptToCloseProcesses"].Split(',');
                var displayNames = session["PromptToCloseDisplayNames"].Split(',');

                if (processes.Length != displayNames.Length)
                {
                    session.Log(@"Please check that 'PromptToCloseProcesses' and 'PromptToCloseDisplayNames' exist and have same number of items.");
                    MessageBox.Show("安装过程中出现错误，请关闭杀毒软件后再次尝试安装", "SumscopeAddIn");
                    return ActionResult.Failure;
                }

                for (var i = 0; i < processes.Length; i++)
                {
                    session.Log("Prompting process {0} with name {1} to close.", processes[i], displayNames[i]);
                    using (var prompt = new PromptCloseApplication(productName, processes[i], displayNames[i]))
                    {
                        if (!prompt.Prompt())
                        {
                            MessageBox.Show("安装过程中出现错误，请关闭杀毒软件后再次尝试安装", "SumscopeAddIn");
                            return ActionResult.Failure;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                session.Log("Missing properties or wrong values. Please check that 'PromptToCloseProcesses' and 'PromptToCloseDisplayNames' exist and have same number of items. \nException:" + ex.Message);
                MessageBox.Show("安装过程中出现错误，请关闭杀毒软件后再次尝试安装", "SumscopeAddIn");
                return ActionResult.Failure;
            }

            session.Log("End PromptToCloseApplications");
            return ActionResult.Success;
        }
        #endregion

        #region GetAddInName
        private static string GetAddInName(string szXll32Name, string szXll64Name, string szOfficeVersionKey, double nVersion)
        {
            Console.WriteLine("检测Office位数 {0}...", nVersion);
            var officeBitness = GetOfficeBitness(szOfficeVersionKey, nVersion);
            pulicSession.Log("GetOfficeBitness值为：{0}", officeBitness);
            switch (officeBitness)
            {
                case OfficeBitness.X86:
                    pulicSession.Log("Office 32 bits.");
                    pulicSession.Log("即将返回：" + szXll32Name);
                    return szXll32Name;

                case OfficeBitness.X64:
                    pulicSession.Log("Office 64 bits.");
                    return szXll64Name;
                default:
                    throw new InvalidOperationException("异常：未能检测出Office版本位数 " + nVersion);
            }
        }

        #endregion

        #region GetOfficeBitness
        private static OfficeBitness GetOfficeBitness(string szOfficeVersionKey, double nVersion)
        {
            // before office 2010, no 64 bits version of office exists. also only 32 bits can be installed on 32 bits systems.
            if (nVersion < 14 || !Environment.Is64BitOperatingSystem)
            {
                return OfficeBitness.X86;
            }

            // Check the ClickToRun registry (x86+x64). Both must be checked.
            // http://msdn.microsoft.com/en-us/library/office/ff864733(v=office.15).aspx
            RegistryKey clickToRunRegKey86 = RegistryKey
                .OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Registry32)
                .OpenSubKey(@"Software\Microsoft\Office\" + szOfficeVersionKey + @"\ClickToRun\Configuration", false);
            Console.WriteLine("Office bitness using clicktorun x86 office installation: {0}present", clickToRunRegKey86 == null ? "not " : "");
            RegistryKey clickToRunRegKey64 = RegistryKey
                .OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Registry64)
                .OpenSubKey(@"Software\Microsoft\Office\" + szOfficeVersionKey + @"\ClickToRun\Configuration", false);
            Console.WriteLine("Office bitness using clicktorun x64 office installation: {0}present", clickToRunRegKey64 == null ? "not " : "");


            // Check the Outlook\Bitness registry key
            // Using a registry key of outlook to determine the bitness of office may look like weird but that's the reality.
            // http://stackoverflow.com/questions/2203980/detect-whether-office-2010-is-32bit-or-64bit-via-the-registry

            // Note about upgrading office with "keep previous version" option:
            // Only one version of Outlook can be installed at a time. However, we can have several excel, word, etc versions at the same time.
            // One of the Outlook registry key is removed when upgrading Office. Thus the bitness is not found, resulting in the setup to fail.
            // Checking both x86/64 keys for office bitness seems to do the job.

            // Another alternative might be to check the bitness of any version of Office. It seems that you can't install 32bits and 64bits version
            // of office side-by-side (https://msdn.microsoft.com/en-us/library/ee691831.aspx#Anchor_6, https://technet.microsoft.com/en-us/library/ee681792.aspx)

            RegistryKey outlookRegKey86 =
                RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Registry32)
                    .OpenSubKey(@"Software\Microsoft\Office\" + szOfficeVersionKey + @"\Outlook", false);
            Console.WriteLine("Office bitness using std x86 office installation: {0}present", outlookRegKey86 == null ? "not " : "");
            RegistryKey outlookRegKey64 =
                RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Registry64)
                    .OpenSubKey(@"Software\Microsoft\Office\" + szOfficeVersionKey + @"\Outlook", false);
            Console.WriteLine("Office bitness using std x64 office installation: {0}present", outlookRegKey64 == null ? "not " : "");


            // First check clicktorun (skip if not defined), new deployment tool from microsoft
            var bitnessRegKey = clickToRunRegKey86 ?? clickToRunRegKey64;
            if (bitnessRegKey != null)
            {
                switch ((bitnessRegKey.GetValue("Platform") ?? "").ToString())
                {
                    case "x64":
                        return OfficeBitness.X64;
                    case "x86":
                        return OfficeBitness.X86;
                }
            }

            // Then check outlook bitness registry key
            var outlookRegKeys = new List<RegistryKey> { outlookRegKey86, outlookRegKey64 };
            foreach (var outlookRegKey in outlookRegKeys.Where(x => x != null))
            {
                object oBitValue = outlookRegKey.GetValue("Bitness");
                if (oBitValue != null)
                {
                    switch (oBitValue.ToString())
                    {
                        case "x64":
                            return OfficeBitness.X64;
                        case "": // Empty key means x86 for older install of office.
                        case "x86":
                            return OfficeBitness.X86;
                    }
                }
            }

            // If not found, then unknown
            return OfficeBitness.Unknown;
        }
        #endregion

        #region IsOfficeExcelInstalled 判断指定版本的excel是否安装 因为可能会安装多个版本的excel 让每一个版本的excel都加载上 excelAddin
        private static bool IsOfficeExcelInstalled(string excelBaseKey)
        {
            // Check both x86 and x64 registry
            var hklmRoot = new List<RegistryKey>
            {
                RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Registry64),
                RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Registry32)
            };

            /*
             * Here, we check if excel is trully installed on the system by checking Office installation root + application name.
             * HKLM\Software\Microsoft\Office\x.x\Excel\InstallRoot | Path
             */

            var excelInstallRootKey = excelBaseKey + @"\InstallRoot";
            foreach (var root in hklmRoot)
            {
                var installRootKey = root.OpenSubKey(excelInstallRootKey, false);
                if (installRootKey == null)
                {
                    continue;
                }

                var pathKey = installRootKey.GetValue("Path") as string;
                if (string.IsNullOrEmpty(pathKey))
                {
                    continue;
                }

                try
                {
                    var excelApplicationPath = Path.Combine(pathKey, "excel.exe");
                    if (File.Exists(excelApplicationPath))
                    {
                        return true;
                    }
                }
                catch (ArgumentException ex)
                {
                    // if the registry key is corrupted (Path.Combine call), we don't want to throw. but log it just in case.
                    Console.WriteLine("IsOfficeExcelInstalled failed due to invalid value in registry key {0}. Consider Microsoft Office Excel not installed for this version. Exception: {1}", excelInstallRootKey, ex);
                }
            }

            return false;
        }
        #endregion

        #endregion

        private enum OfficeBitness
        {
            Unknown,
            X86,
            X64
        }
    }
}
