using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Formatters.Binary;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace LazyKMS
{
    public partial class MainForm : Form
    {
        public MainForm()
        {
            InitializeComponent();
        }

        private void RunAction(int i)
        {
            // Don't use FormProvider since we want to have control over
            // some variables and initialize new instance of the form.
            ProcessForm procf = new ProcessForm();
            procf.Action = i;
            procf.Show();
        }

        private void MainForm_Load(object sender, EventArgs e)
        {
            // Because winforms bad
            comboBox2.DropDownStyle = ComboBoxStyle.DropDownList;
            
            // Load settings
            SetSettings();

            // Windows Product Information
            string wininfo = $"Full name: {Lazy.GetWindowsFull()}\nProduct ID: {Lazy.GetWindowsProductId()}\nEdition: {Lazy.GetWindowsEdition()}";
            richTextBox2.Text = wininfo;
        }

        private void SetSettings()
        {
            SettingsHelper.settings.kmsserver = comboBox1.Text;

            SettingsHelper.settings.forcekey = checkBox1.Checked;
            SettingsHelper.settings.key = textBox1.Text;
        }

        private void button5_Click(object sender, EventArgs e)
        {
            SetSettings();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            RunAction(0);
            this.Hide();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            RunAction(1);
            this.Hide();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            RunAction(2);
            this.Hide();
        }

        /*private void button6_Click(object sender, EventArgs e)
        {
            Stream stream;
            SaveFileDialog savediag = new SaveFileDialog();

            savediag.Filter = "JSON config (*.json)|*.json";
            savediag.FilterIndex = 2;
            savediag.RestoreDirectory = true;

            if (savediag.ShowDialog() == DialogResult.OK)
            {
                if ((stream = savediag.OpenFile()) != null)
                {          
                    string json = JsonConvert.SerializeObject(SettingsHelper.settings);

                    StreamWriter sw = new StreamWriter(stream);
                    sw.Write(json);
                    
                    sw.Close();
                    stream.Close();
                }
            }
        }

        private void button7_Click(object sender, EventArgs e)
        {
            Stream stream;
            OpenFileDialog opendiag = new OpenFileDialog();

            opendiag.Filter = "JSON config (*.json)|*.json";
            opendiag.FilterIndex = 2;
            opendiag.RestoreDirectory = true;

            if (opendiag.ShowDialog() == DialogResult.OK)
            {
                if ((stream = opendiag.OpenFile()) != null)
                {
                    StreamReader sr = new StreamReader(stream);
                    string json = sr.ReadToEnd();

                    Settings s = JsonConvert.DeserializeObject<Settings>(json);

                    comboBox1.Text = s.kmsserver;
                    checkBox1.Checked = s.forcekey;
                    textBox1.Text = s.key;

                    sr.Close();
                    stream.Close();
                }
            }
        }*/

        private void SetOfficeVersion()
        {
            SettingsHelper.settings.officever = comboBox2.Text;
        }

        private void button8_Click(object sender, EventArgs e)
        {
            SetOfficeVersion();
            MessageBox.Show("Please read carefully!\n\nIt is really hard to detect where all of Office versions are installed since M$ is moving them like shit and registry keys can be completely same for different versions, so this program uses default install paths for activation.\n\nMake sure you have installed office in default installation directory!", "Version Information", MessageBoxButtons.OK, MessageBoxIcon.Information);

            RunAction(3);
            this.Hide();
        }

        private void button8_Click_1(object sender, EventArgs e)
        {
            SetOfficeVersion();
            RunAction(3);
            this.Hide();
        }

        private void button7_Click(object sender, EventArgs e)
        {
            SetOfficeVersion();
            RunAction(4);
            this.Hide();
        }

        private void button6_Click(object sender, EventArgs e)
        {
            SetOfficeVersion();
            RunAction(5);
            this.Hide();
        }
    }
}
