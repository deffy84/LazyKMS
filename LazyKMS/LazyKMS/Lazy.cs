using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace LazyKMS
{
    class Lazy
    {
        const string currentVersion = @"HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion";
        const string sys32 = @"C:\Windows\System32\";

        /// <summary>
        /// Gets full Windows name as shown in System Info
        /// </summary>
        /// <returns>Text (Windows 10 Enterprise LTSC 2019)</returns>
        public static string GetWindowsFull()
        {
            return Registry.GetValue(currentVersion, "ProductName", "Unknown").ToString();
        }

        /// <summary>
        /// Gets Windows product ID
        /// </summary>
        /// <returns>Text (0eb15-9b33a-18992-a7083)</returns>
        public static string GetWindowsProductId()
        {
            return Registry.GetValue(currentVersion, "ProductId", "Unknown").ToString();
        }

        /// <summary>
        /// Gets Windows edition
        /// </summary>
        /// <returns>Text (EnterpriseS)</returns>
        public static string GetWindowsEdition()
        {
            return Registry.GetValue(currentVersion, "EditionId", "Unknown").ToString();
        }

        /// <summary>
        /// Run external process and output it's stdout and stderr
        /// to EventHandler
        /// </summary>
        /// <param name="exe">File path to executable</param>
        /// <param name="args">Process start arguments</param>
        /// <param name="output">EventHandler to receive process output</param>
        public static void RunProcess(string exe, string args, DataReceivedEventHandler output)
        {
            Process process = new Process();
            process.StartInfo.FileName = exe;
            process.StartInfo.Arguments = args;
            process.StartInfo.UseShellExecute = false;
            process.StartInfo.RedirectStandardOutput = true;
            process.StartInfo.RedirectStandardError = true;
            process.StartInfo.CreateNoWindow = true;
            process.OutputDataReceived += new DataReceivedEventHandler(output);
            process.ErrorDataReceived += new DataReceivedEventHandler(output);
            process.Start();
            process.BeginOutputReadLine();
            process.BeginErrorReadLine();
            process.WaitForExit();
        }

        /// <summary>
        /// Set Windows key (slmgr.vbs /ipk)
        /// </summary>
        /// <param name="key">Key to set</param>
        /// <param name="output">EventHandler to receive slmgr output</param>
        public static void SetKey(string key, DataReceivedEventHandler output)
        {
            RunProcess("cscript", $"{sys32}slmgr.vbs /ipk {key}", output);
        }

        /// <summary>
        /// Set Windows KMS server (slmgr.vbs /skms)
        /// </summary>
        /// <param name="server">Server address</param>
        /// <param name="output">EventHandler to receive slmgr output</param>
        public static void SetServer(string server, DataReceivedEventHandler output)
        {
            RunProcess("cscript", $"{sys32}slmgr.vbs /skms {server}", output);
        }

        /// <summary>
        /// Run Windows activation with current settings (slmgr.vbs /ato)
        /// </summary>
        /// <param name="output">EventHandler to receive slmgr output</param>
        public static void Activate(DataReceivedEventHandler output)
        {
            RunProcess("cscript", $"{sys32}slmgr.vbs /ato", output);
        }

        /// <summary>
        /// Gets Windows volume licensing key
        /// </summary>
        /// <param name="fullname">Full os name (from GetWindowsFull())</param>
        /// <returns></returns>
        public static string GetWinKey(string fullname)
        {
            return KeyList.GetWindowsKey(fullname);
        }


    }
}
