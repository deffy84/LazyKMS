using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.Linq;
using System.Reflection;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace LazyKMS
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            // Check if we are running with Windows in English
            CultureInfo ci = CultureInfo.InstalledUICulture;
            if (ci.Name != "en-US")
            {
                MessageBox.Show("You are not using US English as your main OS language. Program may invalidly report success status, so please check console if any errors really happen!", "Little Problem", MessageBoxButtons.OK, MessageBoxIcon.Warning);

            }

            Application.Run(FormProvider.mainForm);
        }
    }
}
