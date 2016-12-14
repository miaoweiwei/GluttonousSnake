using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Deployment.WindowsInstaller;

namespace InstallerCA
{
    class Parameters
    {
        public string Xll32Name { get; set; }
        public string Xll64Name { get; set; }
        public List<string> SupportedOfficeVersion { get; set; }
        public string InstallDirectory { get; set; }

        public static Parameters ExtractFromSession(Session session)
        {
            session.Log("获取属性值：OFFICEREGKEYS;XLL32;XLL64;AddinFolder");
            string officeRegKeyVersions = session["OFFICEREGKEYS"];
            string xll32Name = session["XLL32"];
            string xll64Name = session["XLL64"];
            string installDirectory = session["AddinFolder"];

            string isCheck = "判断属性值";
            string isCheckOK = "判断属性值OK";

            Parameters parameters = new Parameters();

            #region SupportedOfficeVersion
            session.Log("{0}：OFFICEREGKEYS......", isCheck);
            List<string> officeRegKeyVersionsList = officeRegKeyVersions.Split(',').ToList();
            if (officeRegKeyVersionsList.Count <= 0)
            {
                throw new ArgumentException("异常：输入至少支持一种版本Office的内部版本号，设置属性[OFFICEREGKEYS]");
            }
            parameters.SupportedOfficeVersion = officeRegKeyVersionsList;
            session.Log("{0}：OFFICEREGKEYS={1}", isCheckOK, officeRegKeyVersions);
            #endregion

            #region InstallDirectory
            session.Log("{0}：installDirectory......", isCheck);
            if (string.IsNullOrEmpty(installDirectory))
            {
                throw new ArgumentException("异常：属性[installDirectory]不得为空");
            }
            parameters.InstallDirectory = installDirectory;
            session.Log("{0}：installDirectory={1}", isCheckOK, installDirectory);
            #endregion

            #region Xll32Name
            session.Log("{0}：XLL32......", isCheck);
            if (string.IsNullOrEmpty(xll32Name))
            {
                throw new ArgumentException("异常：属性[XLL32]不得为空");
            }
            parameters.Xll32Name = xll32Name;
            session.Log("{0}：XLL32={1}", isCheckOK, xll32Name);
            #endregion

            #region Xll64Name
            session.Log("{0}：XLL64......", isCheck);
            if (string.IsNullOrEmpty(xll64Name))
            {
                throw new ArgumentException("异常：属性[XLL64]不得为空");
            }
            parameters.Xll64Name = xll64Name;
            session.Log("{0}：XLL64={1}", isCheckOK, xll64Name);
            #endregion

            return parameters;
        }
    }
}
