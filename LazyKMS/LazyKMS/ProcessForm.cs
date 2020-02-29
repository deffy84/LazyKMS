using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace LazyKMS
{
    public partial class ProcessForm : Form
    {
        public int Action;
        private DataReceivedEventHandler _handler;
        private string _procoutput;

        // 0 - nothing, 1 - error, 2 - success
        private int _keyset;
        private int _serverset;
        private int _fullset;
        
        private int _officekey;
        private int _officeserver;

        public ProcessForm()
        {
            InitializeComponent();
        }

        private void SetInfoText(string text)
        {
            label1.Invoke((MethodInvoker)delegate {
                label1.Text = text;
            });
        }

        public void OutputHandler(object sendingProcess, DataReceivedEventArgs outLine)
        {
            richTextBox1.Invoke((MethodInvoker)delegate {
                richTextBox1.Text += outLine.Data + "\n";
            });
            _procoutput += outLine.Data + "\n";
        }

        private void ProcessForm_Load(object sender, EventArgs e)
        {
            _keyset = 0;
            _serverset = 0;
            _fullset = 0;
            _officekey = 0;
            _handler = new DataReceivedEventHandler(OutputHandler);

            switch (Action)
            {
                case 0:
                    SetKey(true);
                    break;

                case 1:
                    SetServer(true);
                    break;

                case 2:
                    Full();
                    break;

                case 3:
                    SetKeyOffice(true);
                    break;

                case 4:
                    SetServerOffice(true);
                    break;

                default:
                    Full();
                    break;
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            FormProvider.mainForm.Show();
            this.Close();
        }

        private void Finish()
        {
            progressBar1.Invoke((MethodInvoker)delegate {
                progressBar1.Style = ProgressBarStyle.Continuous;
                progressBar1.Value = progressBar1.Maximum;
            });

            button1.Invoke((MethodInvoker)delegate {
                button1.Enabled = true;
            });

            label1.Invoke((MethodInvoker)delegate {
                label1.Text = "Finished. See console for details.";
            });
        }

        private void SetKey(bool showend)
        {
            new Thread(() =>
            {
                Thread.CurrentThread.IsBackground = true;

                SetInfoText("Getting Windows information...");
                string full = Lazy.GetWindowsFull();

                SetInfoText("Getting key...");
                string key = Lazy.GetWinKey(full);
                if (SettingsHelper.settings.forcekey)
                {
                    key = SettingsHelper.settings.key;
                }
                if (key == "unknown")
                {
                    _keyset = 1;
                    MessageBox.Show($"Failed to get a key for your Windows edition/version. Please submit an issue report to official GitHub and include screenshot of this Window.\n\nFull name: {full}\nProduct ID: {Lazy.GetWindowsProductId()}\n\nYou can also create a pull request with the right key for your version. If you know the key you can set in manually in settings by checking the force key overwrite option.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }

                SetInfoText("Setting key...");
                Lazy.SetKey(key, _handler);
                if (_procoutput.Contains("Installed product key")
                && _procoutput.Contains("successfully"))
                {
                    _keyset = 2;
                }
                else
                {
                    _keyset = 1;
                }
                _procoutput = "";

                if (showend)
                {
                    if (_keyset == 2)
                    {
                        MessageBox.Show("Key set successfully!", "OK", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    } else
                    {
                        MessageBox.Show("Failed to set key. Please see console output for more information.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                    Finish();
                }

            }).Start();
        }

        private void SetServer(bool showend)
        {
            new Thread(() =>
            {
                Thread.CurrentThread.IsBackground = true;

                SetInfoText("Setting server...");
                Lazy.SetServer(SettingsHelper.settings.kmsserver, _handler);
                if (_procoutput.Contains("Key Management Service machine name set to"))
                {
                    _serverset = 2;
                }
                else
                {
                    _serverset = 1;
                }
                _procoutput = "";

                if (showend)
                {
                    if (_serverset == 2)
                    {
                        MessageBox.Show("Server set successfully!", "OK", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    else
                    {
                        MessageBox.Show("Failed to set server. Please see console output for more information.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                    Finish();
                }
            }).Start();
        }

        private void Full()
        {
            new Thread(() =>
            {
                Thread.CurrentThread.IsBackground = true;

                // multithreading lvl 100
                SetKey(false);
                while (_keyset == 0) { Thread.Sleep(1); }

                SetServer(false);
                while (_serverset == 0) { Thread.Sleep(1); }

                SetInfoText("Activating...");
                Lazy.Activate(_handler);
                if (_procoutput.Contains("Product activated successfully"))
                {
                    _fullset = 2;
                } else
                {
                    _fullset = 1;
                }
                _procoutput = "";

                if (_fullset == 2)
                {
                    MessageBox.Show("Activation successful!", "OK", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    MessageBox.Show("Failed to activate Windows. Please see console output for more information.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                Finish();
            }).Start();
        }

        /*private void OfficeDetect()
        {
            new Thread(() =>
            {
                Thread.CurrentThread.IsBackground = true;

                SetInfoText("Waiting for Word process...");
                Process[] word = Process.GetProcessesByName("WINWORD");
                while (word.Length == 0)
                {
                    word = Process.GetProcessesByName("WINWORD");
                    Thread.Sleep(1000);
                }
                string exepath = word[0].MainModule.FileName;
                // C:\Program Files(x86)\Microsoft Office\root\Office16
                // -> C:\Program Files (x86)\Microsoft Office
                string activpath = Directory.GetParent(Directory.GetParent(Path.GetDirectoryName(exepath)).FullName).FullName; // oof
                string officefol = new DirectoryInfo(Path.GetDirectoryName(exepath)).Name;
                // C:\Program Files (x86)\Microsoft Office
                // -> C:\Program Files (x86)\Microsoft Office\Office16
                activpath += "\\" + officefol + "\\";

                SetInfoText("Getting Office version...");
                Lazy.ActivateOffice(activpath, _handler);
                if (_procoutput.Contains("Key Management Service machine name set to"))
                {
                    _serverset = 2;
                }
                else
                {
                    _serverset = 1;
                }
                _procoutput = "";

                Finish();
            }).Start();
        }*/

        private void SetKeyOffice(bool showend)
        {
            new Thread(() =>
            {
                Thread.CurrentThread.IsBackground = true;

                SetInfoText("Getting Office install directory...");
                string officepath = Lazy.GetOfficeDir(SettingsHelper.settings.officever);

                SetInfoText("Getting key...");
                string key = Lazy.GetOfficeKey(SettingsHelper.settings.officever);

                SetInfoText("Setting key...");
                Lazy.SetKeyOffice(officepath, key, _handler);
                if (_procoutput.Contains("Product key installation successful"))
                {
                    _officekey = 2;
                }
                else
                {
                    _officekey = 1;
                }
                _procoutput = "";

                if (showend)
                {
                    if (_officekey == 2)
                    {
                        MessageBox.Show("Key set successfully!", "OK", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    else
                    {
                        MessageBox.Show("Failed to set key. Please see console output for more information.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                    Finish();
                }

            }).Start();
        }

        private void SetServerOffice(bool showend)
        {
            new Thread(() =>
            {
                Thread.CurrentThread.IsBackground = true;

                SetInfoText("Getting Office install directory...");
                string officepath = Lazy.GetOfficeDir(SettingsHelper.settings.officever);

                SetInfoText("Setting server...");
                Lazy.SetServerOffice(officepath, SettingsHelper.settings.kmsserver, _handler);

                if (_procoutput.Contains("Successfully applied setting"))
                {
                    _officeserver = 2;
                }
                else
                {
                    _officeserver = 1;
                }
                _procoutput = "";

                if (showend)
                {
                    if (_officeserver == 2)
                    {
                        MessageBox.Show("Server set successfully!", "OK", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    else
                    {
                        MessageBox.Show("Failed to set server. Please see console output for more information.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                    Finish();
                }
            }).Start();
        }


    }
}
