using Microsoft.Win32;
using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Timers;
using System.Windows.Forms;
using static LogHelp;

namespace Gama
{
    public partial class Gama : Form
    {
        private OPCData OPCData;
        private static string Company;
        private static string AutoStart;
        public Gama()
        {
            InitializeComponent();
            ShowInTaskbar = false;
        }
        public struct CopyDataStruct
        {
            public IntPtr dwData;
            public int cbData;
            [MarshalAs(UnmanagedType.LPStr)]
            public string lpData;
        }
        private void Gama_Load(object sender, EventArgs e)
        {
            Company = ConfigurationManager.AppSettings["Company"];
            AutoStart = ConfigurationManager.AppSettings["AutoStart"];
            if (AutoStart=="true") {
                //自启动
                RegistryKey registryKey = Registry.CurrentUser.OpenSubKey
                     ("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Run", true);
                registryKey.SetValue("分析仪数据检测终端", Application.ExecutablePath);
            }
            OPCData = new OPCData();
            Thread OPC = new Thread(OPCData.GetLocalServer)
            {
                IsBackground = true
            };
            OPC.Start();
            ///重启线程
            Thread Restart = new Thread(RestartTimer)
            {
                IsBackground = true
            };
            Restart.Start();
        }

        public const int WM_COPYDATA = 0x004A;

        //通过窗口标题来查找窗口句柄   
        [DllImport("User32.dll", EntryPoint = "FindWindow")]
        private static extern int FindWindow(string lpClassName, string lpWindowName);

        //在DLL库中的发送消息函数  
        [DllImport("User32.dll", EntryPoint = "SendMessage")]
        private static extern int SendMessage
            (
            int hWnd,                        // 目标窗口的句柄    
            int Msg,                         // 在这里是WM_COPYDATA  
            int wParam,                      // 第一个消息参数  
            ref CopyDataStruct lParam        // 第二个消息参数  
           );

        //发送消息
        public static void sendMessage(string str)
        {
            try
            {
                string strURL = str;
                CopyDataStruct cds;
                cds.dwData = (IntPtr)1; //这里可以传入一些自定义的数据，但只能是4字节整数        
                cds.lpData = strURL;    //消息字符串  
                cds.cbData = System.Text.Encoding.Default.GetBytes(strURL).Length + 1; //注意，这里的长度是按字节来算的  
                SendMessage(FindWindow(null, Company), WM_COPYDATA, 0, ref cds);       // 窗口标题 
            } 
            catch { }
            }
        //接收消息方法  
        protected override void WndProc(ref System.Windows.Forms.Message e)
        {
            if (e.Msg == WM_COPYDATA)
            {
                CopyDataStruct cds = (CopyDataStruct)e.GetLParam(typeof(CopyDataStruct));
                switch (cds.lpData.ToString()) {
                    case "close":
                        OPCData.opc_connected = false;
                        Application.Exit();
                        break;
                }
               
            }
            base.WndProc(ref e);
        }
        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (OPCData.opc_connected) {
                e.Cancel = true; ;
            }
        }
        public void RestartTimer()
        {
            System.Timers.Timer RestartTimer = new System.Timers.Timer(1000 * 60 * 60);
            RestartTimer.Elapsed += new ElapsedEventHandler(Restart);
            RestartTimer.AutoReset = true;
            RestartTimer.Enabled = true;
            /// <summary>
            /// 定时重启任务
            /// </summary>
            /// <param name="source"></param>
            /// <param name="e"></param>
            void Restart(object source, System.Timers.ElapsedEventArgs e)
            {
                if (DateTime.Now.Hour.ToString() == "0"&&AutoStart == "true")
                {
                    OPCData.opc_connected = false;
                    //开启新的实例
                    Process.Start(Application.ExecutablePath);
                    //关闭当前实例
                    Application.Exit();
                }
            }
        }
    }
}
