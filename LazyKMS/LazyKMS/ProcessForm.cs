using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
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
            _handler = new DataReceivedEventHandler(OutputHandler);

            switch (Action)
            {
                case 0:
                    OnlyKey();
                    break;

                case 1:
                    OnlyServer();
                    break;

                case 2:
                    Full();
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

        private void OnlyKey()
        {
            SetKey(true);
        }

        private void OnlyServer()
        {
            SetServer(true);
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
    }
}
