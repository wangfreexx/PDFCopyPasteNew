using System;
using System.Configuration;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Windows.Forms;

namespace PDFCopyPaste
{
    class Form1 : Form
    {
        [System.Runtime.InteropServices.DllImport("user32")]
        private static extern IntPtr SetClipboardViewer(IntPtr hwnd);
        [System.Runtime.InteropServices.DllImport("user32")]
        private static extern IntPtr ChangeClipboardChain(IntPtr hwnd, IntPtr hWndNext);
        [System.Runtime.InteropServices.DllImport("user32")]
        private static extern int SendMessage(IntPtr hwnd, int wMsg, IntPtr wParam, IntPtr lParam);
        [System.Runtime.InteropServices.DllImport("user32")]
        private static extern IntPtr GetClipboardOwner();
        [System.Runtime.InteropServices.DllImport("user32")]
        private static extern int GetWindowThreadProcessId(IntPtr handle, out int processId);
        [System.Runtime.InteropServices.DllImport("user32")]
        private static extern bool CloseHandle(IntPtr handle);
        [System.Runtime.InteropServices.DllImport("kernel32")]
        public static extern int OpenProcess(int dwDesiredAccess, bool bInheritHandle, int dwProcessId);
        [DllImport("Psapi.dll", EntryPoint = "GetModuleFileNameEx")]
        public static extern uint GetModuleFileNameEx(int handle, int hModule, [Out] StringBuilder lpszFileName, uint nSize);




        const int WM_DRAWCLIPBOARD = 0x308;
        private GroupBox groupBox1;
        private NotifyIcon notifyIcon1;
        private System.ComponentModel.IContainer components;
        private ContextMenuStrip contextMenuStrip1;
        private ToolStripMenuItem 退出ToolStripMenuItem;
        private CheckBox checkBox5;
        private CheckBox checkBox4;
        private CheckBox checkBox1;
        private CheckBox checkBox3;
        private CheckBox checkBox2;
        private TextBox textBox1;
        private Label label1;
        const int WM_CHANGECBCHAIN = 0x30D;
        string[] need_deal_soft;

        public Form1()
        {
            Load += Form1_Load;
            Closed += Form1_Closed;
            Shown += Forms1_Shown;
            //WindowState = FormWindowState.Minimized;
            InitializeComponent();


            this.MaximizeBox = false;
            if (checkBox4.Checked == true)
            {
                Width = 230;
            }
            else if (checkBox4.Checked == false)
            {
                Width = 459;
            }
            //False
            try
            {
                checkBox1.Checked = bool.Parse(ConfigurationManager.AppSettings["check1"]);
                checkBox2.Checked = bool.Parse(ConfigurationManager.AppSettings["check2"]);
                checkBox3.Checked = bool.Parse(ConfigurationManager.AppSettings["check3"]);
                checkBox4.Checked = bool.Parse(ConfigurationManager.AppSettings["check4"]);
                checkBox5.Checked = bool.Parse(ConfigurationManager.AppSettings["check5"]);
                textBox1.Text = ConfigurationManager.AppSettings["text1"];
            }
            catch
            {
                checkBox1.Checked = true;
                checkBox2.Checked = true;
                checkBox3.Checked = true;
                checkBox4.Checked = true;
                checkBox5.Checked = true;
                textBox1.Text = "PDF|pdf|CAJ|caj|DocBox";
            }



            need_deal_soft = textBox1.Text.Split('|');



        }

        private void Forms1_Shown(object sender, EventArgs e)
        {
            //Visible = false;
        }

        private string ProcessText(string str)
        {
            //string str = str1;
            str = str.Replace("", "");

            // 全角转半角
            if (checkBox3.Checked == true)
            {
                str = str.Normalize(NormalizationForm.FormKC);
            }
            //合并换行
            if (checkBox1.Checked == true || checkBox2.Checked == true)
            {
                for (var counter = 0; counter < str.Length - 1; counter++)
                {
                    if (checkBox1.Checked == true)
                    {
                        //合并换行
                        if (str[counter + 1].ToString() == "\r" || str[counter + 1].ToString() == "\r\n" || str[counter + 1].ToString() == "\n")
                        {
                            //如果检测到句号结尾,则不去掉换行
                            if (str[counter].ToString() == "." || str[counter].ToString() == "。") continue;

                            //去除换行
                            try
                            {
                                str = str.Remove(counter + 1, 2);
                            }
                            catch
                            {
                                str = str.Remove(counter + 1, 1);
                            }


                            //判断英文单词或,结尾,则加一个空格
                            if (Regex.IsMatch(str[counter].ToString(), "[a-zA-Z]") || str[counter].ToString() == ",")
                                str = str.Insert(counter + 1, " ");

                            //判断"-"结尾,且前一个字符为英文单词,则去除"-"
                            if (str[counter].ToString() == "-" && Regex.IsMatch(str[counter - 1].ToString(), "[a-zA-Z]"))
                                str = str.Remove(counter, 1);
                        }
                    }
                    //检测到中文时去除空格
                    if (checkBox2.Checked == true && Regex.IsMatch(str, @"[\u4e00-\u9fa5]") && str[counter].ToString() == " ")
                    {
                        str = str.Remove(counter, 1);
                    }
                }
            }
            return str;
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            this.ShowInTaskbar = true;
            //托盘区图标隐藏
            notifyIcon1.Visible = false;
            //获得观察链中下一个窗口句柄
            NextClipHwnd = SetClipboardViewer(this.Handle);
        }

        private static object lck = new object();
        const int PROCESS_QUERY_INFORMATION = 0x0400;
        const int PROCESS_VM_READ = 0x0010;

        protected override void WndProc(ref System.Windows.Forms.Message m)
        {
            switch (m.Msg)
            {
                case WM_DRAWCLIPBOARD:
                    try
                    {
                        //检测文本
                        IDataObject obj = Clipboard.GetDataObject();
                        string[] formats = obj.GetFormats();
                        var owner = GetClipboardOwner();
                        int pid;
                        GetWindowThreadProcessId(owner, out pid);
                        int hProcess = OpenProcess(PROCESS_QUERY_INFORMATION | PROCESS_VM_READ, false, pid);
                        StringBuilder buffer = new StringBuilder();
                        buffer.EnsureCapacity(100);
                        GetModuleFileNameEx(hProcess, 0, buffer, 100);
                        if (obj.GetDataPresent("System.String"))
                        {
                            if (checkBox4.Checked == true)
                            {
                                string str = (string)obj.GetData("System.String");
                                string res = ProcessText(str);
                                if (res != str)
                                {
                                    Thread.Sleep(200);
                                    Clipboard.SetDataObject(res, true, 10000, 100);
                                    if (checkBox5.Checked == true)
                                    {
                                        Information info = new Information();
                                        info.Show();
                                    }
                                }

                            }
                            else
                            {
                                int needsoftfalg = 0;
                                if (need_deal_soft.Length != 0)
                                {
                                    foreach (var needsoft in need_deal_soft)
                                    {
                                        if ((buffer.ToString().IndexOf(needsoft) != -1))
                                        {
                                            needsoftfalg = 1;
                                            break;
                                        }
                                    }
                                }

                                if (needsoftfalg == 1)
                                {
                                    string str = (string)obj.GetData("System.String");
                                    string res = ProcessText(str);
                                    if (res != str)
                                    {
                                        Thread.Sleep(200);
                                        Clipboard.SetDataObject(res, true, 10000, 100);
                                        if (checkBox5.Checked == true)
                                        {
                                            Information info = new Information();
                                            info.Show();
                                        }
                                    }
                                }

                            }

                        }
                        else if (owner == IntPtr.Zero && obj.GetDataPresent("System.String"))
                        {
                            string str = (string)obj.GetData("System.String");
                            string res = Regex.Replace(str, "\r\n", "\n");
                            res = Regex.Replace(res, " +$", "", RegexOptions.Multiline);
                            if (res != str)
                            {
                                Thread.Sleep(200);
                                Clipboard.SetDataObject(res, true, 10000, 100);
                                if (checkBox5.Checked == true)
                                {
                                    Information info = new Information();
                                    info.Show();
                                }
                            }

                        }

                    }
                    catch (Exception e)
                    {
                        MessageBox.Show($"{e.GetType()}: {e.Message}\nAborting...", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        Application.Exit();
                    }
                    //将WM_DRAWCLIPBOARD消息传递到下一个观察链中的窗口
                    SendMessage(NextClipHwnd, m.Msg, m.WParam, m.LParam);
                    break;
                default:
                    if (m.Msg == Program.WM_PROMPT_EXIT)
                    {
                        if (MessageBox.Show("Close PDFCopyPaste?", "Warning", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2)
                            == DialogResult.Yes)
                        {
                            Application.Exit();
                        }
                    }
                    base.WndProc(ref m);
                    break;
            }
        }

        private void Form1_Closed(object sender, System.EventArgs e)
        {
            //保存参数
            CheckBox[] checks = new CheckBox[] { checkBox1, checkBox2, checkBox3, checkBox4, checkBox5 };
            //string checkstr = "check";
            string key, value;
            var configFile = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);
            var settings = configFile.AppSettings.Settings;
            for (var i = 0; i < checks.Length; i++)
            {
                key = "check" + (i+1).ToString();
                if (checks[i].Checked == true)
                {
                    value = "True";
                }
                else
                {
                    value = "False";
                }


                if (settings[key] == null)
                {
                    settings.Add(key, value);
                }
                else
                {
                    settings[key].Value = value;
                }
            }
            if (textBox1.Text == String.Empty|| settings["text1"] == null)
            {
                settings.Add("text1", "PDF|pdf|CAJ|caj|DocBox");
            }
            else if(textBox1.Text != String.Empty)
            {
                settings["text1"].Value = textBox1.Text;
            }

            try
            {
                configFile.Save(ConfigurationSaveMode.Modified);
                ConfigurationManager.RefreshSection(configFile.AppSettings.SectionInformation.Name);
            }
            catch (ConfigurationErrorsException)
            {
                Console.WriteLine("Error writing app settings");
            }



            //从观察链中删除本观察窗口（第一个参数：将要删除的窗口的句柄；第二个参数：观察链中下一个窗口的句柄 ）
            ChangeClipboardChain(this.Handle, NextClipHwnd);
            //将变动消息WM_CHANGECBCHAIN消息传递到下一个观察链中的窗口
            SendMessage(NextClipHwnd, WM_CHANGECBCHAIN, this.Handle, NextClipHwnd);
        }

        IntPtr NextClipHwnd;

        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.checkBox5 = new System.Windows.Forms.CheckBox();
            this.checkBox4 = new System.Windows.Forms.CheckBox();
            this.checkBox1 = new System.Windows.Forms.CheckBox();
            this.checkBox3 = new System.Windows.Forms.CheckBox();
            this.checkBox2 = new System.Windows.Forms.CheckBox();
            this.notifyIcon1 = new System.Windows.Forms.NotifyIcon(this.components);
            this.contextMenuStrip1 = new System.Windows.Forms.ContextMenuStrip(this.components);
            this.退出ToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.groupBox1.SuspendLayout();
            this.contextMenuStrip1.SuspendLayout();
            this.SuspendLayout();
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.checkBox5);
            this.groupBox1.Controls.Add(this.checkBox4);
            this.groupBox1.Controls.Add(this.checkBox1);
            this.groupBox1.Controls.Add(this.checkBox3);
            this.groupBox1.Controls.Add(this.checkBox2);
            this.groupBox1.Location = new System.Drawing.Point(12, 3);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(194, 223);
            this.groupBox1.TabIndex = 3;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "功能";
            // 
            // checkBox5
            // 
            this.checkBox5.AutoSize = true;
            this.checkBox5.Checked = true;
            this.checkBox5.CheckState = System.Windows.Forms.CheckState.Checked;
            this.checkBox5.Font = new System.Drawing.Font("微软雅黑", 15F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.checkBox5.Location = new System.Drawing.Point(17, 174);
            this.checkBox5.Name = "checkBox5";
            this.checkBox5.Size = new System.Drawing.Size(151, 31);
            this.checkBox5.TabIndex = 8;
            this.checkBox5.Text = "提示处理完成";
            this.checkBox5.UseVisualStyleBackColor = true;
            // 
            // checkBox4
            // 
            this.checkBox4.AutoSize = true;
            this.checkBox4.Checked = true;
            this.checkBox4.CheckState = System.Windows.Forms.CheckState.Checked;
            this.checkBox4.Font = new System.Drawing.Font("微软雅黑", 15F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.checkBox4.Location = new System.Drawing.Point(17, 137);
            this.checkBox4.Name = "checkBox4";
            this.checkBox4.Size = new System.Drawing.Size(151, 31);
            this.checkBox4.TabIndex = 7;
            this.checkBox4.Text = "处理所有复制";
            this.checkBox4.UseVisualStyleBackColor = true;
            this.checkBox4.CheckedChanged += new System.EventHandler(this.checkBox4_CheckedChanged);
            // 
            // checkBox1
            // 
            this.checkBox1.AutoSize = true;
            this.checkBox1.Checked = true;
            this.checkBox1.CheckState = System.Windows.Forms.CheckState.Checked;
            this.checkBox1.Font = new System.Drawing.Font("微软雅黑", 15F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.checkBox1.Location = new System.Drawing.Point(17, 27);
            this.checkBox1.Name = "checkBox1";
            this.checkBox1.Size = new System.Drawing.Size(111, 31);
            this.checkBox1.TabIndex = 4;
            this.checkBox1.Text = "合并换行";
            this.checkBox1.UseVisualStyleBackColor = true;
            // 
            // checkBox3
            // 
            this.checkBox3.AutoSize = true;
            this.checkBox3.Checked = true;
            this.checkBox3.CheckState = System.Windows.Forms.CheckState.Checked;
            this.checkBox3.Font = new System.Drawing.Font("微软雅黑", 15F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.checkBox3.Location = new System.Drawing.Point(17, 101);
            this.checkBox3.Name = "checkBox3";
            this.checkBox3.Size = new System.Drawing.Size(131, 31);
            this.checkBox3.TabIndex = 6;
            this.checkBox3.Text = "全角转半角";
            this.checkBox3.UseVisualStyleBackColor = true;
            // 
            // checkBox2
            // 
            this.checkBox2.AutoSize = true;
            this.checkBox2.Checked = true;
            this.checkBox2.CheckState = System.Windows.Forms.CheckState.Checked;
            this.checkBox2.Font = new System.Drawing.Font("微软雅黑", 15F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.checkBox2.Location = new System.Drawing.Point(17, 64);
            this.checkBox2.Name = "checkBox2";
            this.checkBox2.Size = new System.Drawing.Size(111, 31);
            this.checkBox2.TabIndex = 5;
            this.checkBox2.Text = "去除空格";
            this.checkBox2.UseVisualStyleBackColor = true;
            // 
            // notifyIcon1
            // 
            this.notifyIcon1.ContextMenuStrip = this.contextMenuStrip1;
            this.notifyIcon1.Icon = ((System.Drawing.Icon)(resources.GetObject("notifyIcon1.Icon")));
            this.notifyIcon1.Text = "pdf复制";
            this.notifyIcon1.Visible = true;
            this.notifyIcon1.DoubleClick += new System.EventHandler(this.notifyIcon1_DoubleClick);
            // 
            // contextMenuStrip1
            // 
            this.contextMenuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.退出ToolStripMenuItem});
            this.contextMenuStrip1.Name = "contextMenuStrip1";
            this.contextMenuStrip1.Size = new System.Drawing.Size(101, 26);
            // 
            // 退出ToolStripMenuItem
            // 
            this.退出ToolStripMenuItem.Name = "退出ToolStripMenuItem";
            this.退出ToolStripMenuItem.Size = new System.Drawing.Size(100, 22);
            this.退出ToolStripMenuItem.Text = "退出";
            this.退出ToolStripMenuItem.Click += new System.EventHandler(this.退出ToolStripMenuItem_Click);
            // 
            // textBox1
            // 
            this.textBox1.Font = new System.Drawing.Font("微软雅黑", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.textBox1.Location = new System.Drawing.Point(229, 67);
            this.textBox1.Multiline = true;
            this.textBox1.Name = "textBox1";
            this.textBox1.ScrollBars = System.Windows.Forms.ScrollBars.Horizontal;
            this.textBox1.Size = new System.Drawing.Size(202, 154);
            this.textBox1.TabIndex = 4;
            this.textBox1.Text = "PDF|pdf|CAJ|caj|DocBox";
            this.textBox1.TextChanged += new System.EventHandler(this.textBox1_TextChanged);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("微软雅黑", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label1.Location = new System.Drawing.Point(225, 19);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(186, 42);
            this.label1.TabIndex = 5;
            this.label1.Text = "需要处理的软件的进程名\r\n关键字，使用|隔开";
            // 
            // Form1
            // 
            this.ClientSize = new System.Drawing.Size(443, 238);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.textBox1);
            this.Controls.Add(this.groupBox1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "Form1";
            this.Text = "PDF复制助手";
            this.SizeChanged += new System.EventHandler(this.Form1_SizeChanged);
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.contextMenuStrip1.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        private void notifyIcon1_DoubleClick(object sender, EventArgs e)
        {
            if (WindowState == FormWindowState.Minimized)
            {
                //还原窗体显示    
                WindowState = FormWindowState.Normal;
                //激活窗体并给予它焦点
                this.Activate();
                //任务栏区显示图标
                this.ShowInTaskbar = true;
                //托盘区图标隐藏
                notifyIcon1.Visible = false;
                NextClipHwnd = SetClipboardViewer(this.Handle);
            }
        }

        private void Form1_SizeChanged(object sender, EventArgs e)
        {
            if (WindowState == FormWindowState.Minimized)
            {
                //隐藏任务栏区图标
                this.ShowInTaskbar = false;
                //图标显示在托盘区
                notifyIcon1.Visible = true;
                NextClipHwnd = SetClipboardViewer(this.Handle);
            }
        }

        private void 退出ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("是否确认退出程序？", "退出", MessageBoxButtons.OKCancel, MessageBoxIcon.Question) == DialogResult.OK)
            {
                // 关闭所有的线程
                this.Dispose();
                this.Close();
            }
        }

        private void checkBox4_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox4.Checked == true)
            {
                Width = 230;
            }
            else if (checkBox4.Checked == false)
            {
                Width = 459;
            }
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            string temp1 = "PDF|pdf|CAJ|caj|DocBox";
            if (textBox1.Text == String.Empty)
            {
                need_deal_soft = temp1.Split('|');
            }
            else
            {
                need_deal_soft = textBox1.Text.Split('|');
            }
        }
    }

    class Program
    {
        static Mutex m_mutex = new Mutex(false, "{BB20662C-D8CA-4E1B-A1B5-ABDC73D14926}");
        [DllImport("user32")]
        public static extern bool PostMessage(IntPtr hwnd, int msg, IntPtr wparam, IntPtr lparam);
        [DllImport("user32")]
        public static extern int RegisterWindowMessage(string message);
        public static int WM_PROMPT_EXIT = RegisterWindowMessage("WM_PROMPT_EXIT");
        [STAThread]
        static void Main(string[] args)
        {
            if (m_mutex.WaitOne(0, true))
            {
                Application.Run(new Form1());
                m_mutex.ReleaseMutex();
            }
            else
            {
                PostMessage(new IntPtr(0xffff), WM_PROMPT_EXIT, IntPtr.Zero, IntPtr.Zero);
            }
        }
    }
}
