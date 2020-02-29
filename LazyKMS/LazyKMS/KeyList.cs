using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace LazyKMS
{
    class KeyList
    {
        /// <summary>
        /// Get Windows volume key
        /// </summary>
        /// <param name="fullname">Text (Windows 10 Enterprise LTSC 2019)</param>
        /// <returns></returns>
        public static string GetWindowsKey(string fullname)
        {
            Dictionary<string, string> list = new Dictionary<string, string>();

            // https://docs.microsoft.com/en-us/windows-server/get-started/kmsclientkeys
            list.Add("Windows Server Datacenter", "6NMRW-2C8FM-D24W7-TQWMY-CWH2D");
            list.Add("Windows Server Standard", "N2KJX-J94YW-TQVFB-DG9YT-724CC");

            list.Add("Windows Server 2019 Datacenter", "WMDGN-G9PQG-XVVXX-R3X43-63DFG");
            list.Add("Windows Server 2019 Standard", "N69G4-B89J2-4G8F4-WWYCC-J464C");
            list.Add("Windows Server 2019 Essentials", "WVDHN-86M7X-466P6-VHXV7-YY726");
            
            list.Add("Windows Server 2016 Datacenter", "CB7KF-BWN84-R7R2Y-793K2-8XDDG");
            list.Add("Windows Server 2016 Standard", "WC2BQ-8NRM3-FDDYY-2BFGV-KHKQY");
            list.Add("Windows Server 2016 Essentials", "JCKRF-N37P4-C2D82-9YXRT-4M63B");
            
            list.Add("Windows 10 Pro", "W269N-WFGWX-YVC9B-4J6C9-T83GX");
            list.Add("Windows 10 Pro N", "MH37W-N47XK-V7XM9-C7227-GCQG9");
            list.Add("Windows 10 Pro for Workstations", "NRG8B-VKK3Q-CXVCJ-9G2XF-6Q84J");
            list.Add("Windows 10 Pro for Workstations N", "9FNHH-K3HBT-3W4TD-6383H-6XYWF");
            list.Add("Windows 10 Pro Education", "6TP4R-GNPTD-KYYHQ-7B7DP-J447Y");
            list.Add("Windows 10 Pro Education N", "YVWGF-BXNMC-HTQYQ-CPQ99-66QFC");
            list.Add("Windows 10 Education", "NW6C2-QMPVW-D7KKK-3GKT6-VCFB2");
            list.Add("Windows 10 Education N", "2WH4N-8QGBV-H22JP-CT43Q-MDWWJ");
            list.Add("Windows 10 Enterprise", "NPPR9-FWDCX-D2C8J-H872K-2YT43");
            list.Add("Windows 10 Enterprise N", "DPH2V-TTNVB-4X9Q3-TJR4H-KHJW4");
            list.Add("Windows 10 Enterprise G", "YYVX9-NTFWV-6MDM3-9PT4T-4M68B");
            list.Add("Windows 10 Enterprise G N", "44RPN-FTY23-9VTTB-MP9BX-T84FV");
            
            list.Add("Windows 10 Enterprise LTSC 2019", "M7XTQ-FN8P6-TTKYV-9D4CC-J462D");
            list.Add("Windows 10 Enterprise N LTSC 2019", "92NFX-8DJQP-P6BBQ-THF9C-7CG2H");
            
            list.Add("Windows 10 Enterprise LTSB 2016", "DCPHK-NFMTC-H88MJ-PFHPY-QJ4BJ");
            list.Add("Windows 10 Enterprise N LTSB 2016", "QFFDN-GRT3P-VKWWX-X7T3R-8B639");
            
            list.Add("Windows 10 Enterprise 2015 LTSB", "WNMTR-4C88C-JK8YV-HQ7T2-76DF9");
            list.Add("Windows 10 Enterprise 2015 LTSB N", "2F77B-TNFGY-69QQF-B8YKP-D69TJ");
            
            list.Add("Windows 8.1 Pro", "GCRJD-8NW9H-F2CDX-CCM8D-9D6T9");
            list.Add("Windows 8.1 Pro N", "HMCNV-VVBFX-7HMBH-CTY9B-B4FXY");

            if (list.ContainsKey(fullname))
            {
                return list[fullname];
            } else
            {
                return "unknown";
            }
        }

        /// <summary>
        /// Get Office volume key
        /// </summary>
        /// <param name="fullname">Text</param>
        /// <returns></returns>
        public static string GetOfficeKey(string fullname)
        {
            Dictionary<string, string> list = new Dictionary<string, string>();

            // http://wind4.github.io/vlmcsd/
            list.Add("Office Professional Plus 2019", "NMMKJ-6RK4F-KMJVX-8D9MJ-6MWKP");
            list.Add("Office Standard 2019", "6NWWJ-YQWMR-QKGCB-6TMB3-9D9HK");
            
            list.Add("Office Professional Plus 2016", "XQNVK-8JYDB-WJ9W3-YJ8YR-WFG99");
            list.Add("Office Standard 2016", "JNRGM-WHDWX-FJJG3-K47QV-DRTFM");

            if (list.ContainsKey(fullname))
            {
                return list[fullname];
            }
            else
            {
                return "unknown";
            }
        }

        /// <summary>
        /// Get Office script directory
        /// (C:\Program Files (x86)\Microsoft Office\Office16)
        /// </summary>
        /// <param name="fullname">Text</param>
        /// <returns></returns>
        public static string GetOfficeScriptDir(string fullname)
        {
            Dictionary<string, string> list = new Dictionary<string, string>();

            // Make sure it ends with \
            list.Add("Office Professional Plus 2019", @"C:\Program Files (x86)\Microsoft Office\Office16\");
            list.Add("Office Standard 2019", @"C:\Program Files (x86)\Microsoft Office\Office16\");

            list.Add("Office Professional Plus 2016", @"C:\Program Files (x86)\Microsoft Office\Office16\");
            list.Add("Office Standard 2016", @"C:\Program Files (x86)\Microsoft Office\Office16\");

            if (list.ContainsKey(fullname))
            {
                return list[fullname];
            }
            else
            {
                return "unknown";
            }
        }
    }
}
