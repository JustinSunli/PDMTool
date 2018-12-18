using System;
using System.Windows.Forms;
using System.Diagnostics;
using System.IO;
using EPDM.Interop.epdm;

namespace TestTool
{
    public partial class Form1 : Form
    {
        private string _ServerExePath;
        private string _ClientExePath;
        private string _VaultName
        {
            get
            {
                return this.textBox1.Text;
            }
        }
        private string _UserName
        {
            get
            {
                return this.textBox2.Text;
            }
        }
        private string _Password
        {
            get
            {
                return this.textBox3.Text;
            }
        }

        private string _Flag
        {
            get
            {
                return this.comboBox1.Text;
            }
        }

        public Form1()
        {
            InitializeComponent();
            
            string thisExePath = Process.GetCurrentProcess().MainModule.FileName;
            string thisExeFolder = Path.GetDirectoryName(thisExePath);
            _ServerExePath = Path.Combine(thisExeFolder, "PDMConnectionExeServer.exe");
            _ClientExePath = Path.Combine(thisExeFolder, "PDMConnectionExeClient.exe");

            this.comboBox1.Text = "1";
        }

        //测试连接
        private void button1_Click(object sender, EventArgs e)
        {
            string error;
            if (!checkInput(out error))
            {
                MessageBox.Show(error);
                return;
            }
            
            string args = combineArgs("-CheckConnection", _VaultName, _UserName, _Password);

            int ret = startProcessAndGetExitCode(_ServerExePath, args);
            string msg = args + "\n返回值 ： " + ret;
            MessageBox.Show(msg);
        }

        //测试检出
        private void button2_Click(object sender, EventArgs e)
        {
            string error;
            if (!checkInput(out error))
            {
                MessageBox.Show(error);
                return;
            }

            string files = "";
            if (string.IsNullOrEmpty(this.textBox4.Text))
            {
                string vaultRootPath = getVaultRootPath();

                OpenFileDialog ofd = new OpenFileDialog();
                ofd.Multiselect = true;
                if (ofd.ShowDialog() != DialogResult.OK || ofd.FileNames.Length == 0)
                    return;

                for (int i = 0; i < ofd.FileNames.Length; ++i)
                {
                    string filePath = ofd.FileNames[i].ToLower();
                    string fileRelativePath = filePath.Replace(vaultRootPath, "");

                    files += fileRelativePath;
                    if (i != ofd.FileNames.Length - 1)
                        files += '|';
                }
            }
            else
                files = this.textBox4.Text;
            
            string args = combineArgs("-CheckOut", _VaultName, _UserName, _Password, _Flag, files);

            int ret = startProcessAndGetExitCode(_ServerExePath, args);
            string msg = args + "\n返回值 ： " + ret;
            MessageBox.Show(msg);
        }

        //测试检入
        private void button3_Click(object sender, EventArgs e)
        {
            string error;
            if (!checkInput(out error))
            {
                MessageBox.Show(error);
                return;
            }

            string files = "";
            if (string.IsNullOrEmpty(this.textBox4.Text))
            {
                OpenFileDialog ofd = new OpenFileDialog();
                ofd.Multiselect = true;
                if (ofd.ShowDialog() != DialogResult.OK || ofd.FileNames.Length == 0)
                    return;

                for (int i = 0; i < ofd.FileNames.Length; ++i)
                {
                    files += ofd.FileNames[i];
                    if (i != ofd.FileNames.Length - 1)
                        files += '|';
                }
            }
            else
                files = this.textBox4.Text;

            string args = combineArgs("-CheckIn", _VaultName, _UserName, _Password, _Flag, files);

            int ret = startProcessAndGetExitCode(_ServerExePath, args);
            string msg = args + "\n返回值 ： " + ret;
            MessageBox.Show(msg);
        }

        //测试查询
        private void button4_Click(object sender, EventArgs e)
        {
            string error;
            if (!checkInput(out error))
            {
                MessageBox.Show(error);
                return;
            }

            string fileName = "";
            if (string.IsNullOrEmpty(this.textBox4.Text))
            {
                string vaultRootPath = getVaultRootPath();

                OpenFileDialog ofd = new OpenFileDialog();
                ofd.Multiselect = false;
                if (ofd.ShowDialog() != DialogResult.OK || ofd.FileNames.Length == 0)
                    return;
                fileName = ofd.FileName.ToLower();
                fileName = fileName.Replace(vaultRootPath, "");
            }
            else
                fileName = this.textBox4.Text;

            string args = combineArgs("-Query", _VaultName, _UserName, _Password, fileName, this.textBox5.Text);

            int ret = startProcessAndGetExitCode(_ServerExePath, args);
            string msg = args + "\n返回值 ： " + ret;
            MessageBox.Show(msg);
        }
        
        //测试刷新
        private void button5_Click(object sender, EventArgs e)
        {
            string error;
            if (!checkInput(out error))
            {
                MessageBox.Show(error);
                return;
            }

            string files = "";
            if (string.IsNullOrEmpty(this.textBox6.Text))
            {
                string vaultRootPath = getVaultRootPath();

                OpenFileDialog ofd = new OpenFileDialog();
                ofd.Multiselect = true;
                if (ofd.ShowDialog() != DialogResult.OK || ofd.FileNames.Length == 0)
                    return;

                for (int i = 0; i < ofd.FileNames.Length; ++i)
                {
                    string filePath = ofd.FileNames[i].ToLower();
                    string fileRelativePath = filePath.Replace(vaultRootPath, "");

                    files += fileRelativePath;
                    if (i != ofd.FileNames.Length - 1)
                        files += '|';
                }
            }
            else
                files = this.textBox6.Text;

            string args = combineArgs("-Refresh", _VaultName, _UserName, _Password, files);

            int ret = startProcessAndGetExitCode(_ClientExePath, args);
            string msg = args + "\n返回值 ： " + ret;
            MessageBox.Show(msg);
        }
        
        private void button6_Click(object sender, EventArgs e)
        {
            string error;
            if (!checkInput(out error))
            {
                MessageBox.Show(error);
                return;
            }

            if (string.IsNullOrEmpty(this.textBox6.Text))
            {
                error = this.label7.Text + "为空";
                MessageBox.Show(error);
                return;
            }

            string args = combineArgs("-GetVaultPath", _VaultName, _UserName, _Password, this.textBox6.Text);
            int ret = startProcessAndGetExitCode(_ClientExePath, args);
            string msg = args + "\n返回值 ： " + ret;
            MessageBox.Show(msg);
        }

        private bool checkInput(out string error)
        {
            if (string.IsNullOrEmpty(_VaultName))
            {
                error = this.label1.Text + "为空";
                return false;
            }

            if (string.IsNullOrEmpty(_UserName))
            {
                error = this.label2.Text + "为空";
                return false;
            }

            if (string.IsNullOrEmpty(_Password))
            {
                error = this.label2.Text + "为空";
                return false;
            }

            error = null;
            return true;
        }

        private string combineArgs(params string[] args)
        {
            string ret = "";
            foreach (string arg in args)
            {
                ret += '"' + arg + '"' + " ";
            }
            ret = ret.TrimEnd();

            return ret;
        }

        private int startProcessAndGetExitCode(string exePath, string args)
        {
            Process p = new Process();
            p.StartInfo.FileName = exePath;
            p.StartInfo.UseShellExecute = false;
            p.StartInfo.Arguments = args;
            p.Start();
            p.WaitForExit();
            return p.ExitCode;
        }

        private string getVaultRootPath()
        {
            string path = null;
            try
            {
                IEdmVault5 vault = new EdmVault5();
                vault.Login(_UserName, _Password, _VaultName);
                path = vault.RootFolderPath;
                path = vault.RootFolderPath.ToLower();
                if (path.LastIndexOf(Path.DirectorySeparatorChar) != path.Length - 1)
                    path += Path.DirectorySeparatorChar;
            }
            catch (Exception ex)
            {
                MessageBox.Show("库登录失败！");
            }

            return path;
        }
    }
}
