using OPCAutomation;
using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Gama.mysqlHelper;
using static LogHelp;
using System.Windows.Forms;

namespace Gama
{
    public class OPCData
    {
        private readonly string Company;
        private OPCServer KepServer;
        private OPCGroups KepGroups;
        private OPCGroup KepGroup;
        private OPCItems KepItems;
        private object serverList;
        private IPHostEntry host;
        private ConcurrentQueue<List<string>> Queue;//数据订阅队列
        private readonly List<string> list;
        private readonly MySqlHelper MySqlHelpers;
        private bool IsCreatServer = true;
        public bool opc_connected = true;
        public OPCData()
        {
            Company = ConfigurationManager.AppSettings["Company"];
            MySqlHelpers = new MySqlHelper();
            Queue = new ConcurrentQueue<List<string>>();
            list = new List<string>() { "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0" };
            Thread upload = new Thread(GamaUpload)
            {
                IsBackground = true
            };
            upload.Start();
        }
        /// <summary>
        /// 获取本地的OPC服务器
        /// </summary>
        public void GetLocalServer()
        {
            host = Dns.GetHostEntry("127.0.0.1");
            var strHostName = host.HostName;
            try
            {
                while (IsCreatServer)
                {
                    try
                    {
                        KepServer = new OPCServer();
                        IsCreatServer = false;
                    }
                    catch (Exception w)
                    {
                        Gama.sendMessage("初始错误");
                        Thread.Sleep(5000);
                    }
                }
                serverList = KepServer.GetOPCServers(strHostName);
                ConnServe(((Array)serverList).GetValue(1).ToString());
            }
            catch
            {
                Gama.sendMessage("服务错误");
                Thread.Sleep(5000);
            }
        }
        /// <summary>
        /// "连接"
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ConnServe(string cmbServerName)
        {
            try
            {
                if (!ConnectRemoteServer("127.0.0.1", cmbServerName))
                {
                    Gama.sendMessage("连接错误");
                    Thread.Sleep(5000);
                    GetLocalServer();
                }
                else
                {
                    if (!CreateGroup())
                    {
                        Gama.sendMessage("组错误");
                    }
                    else
                    {
                        Gama.sendMessage("运行中");
                    }
                }
            }
            catch
            {
                Gama.sendMessage("未知错误");
                Thread.Sleep(5000);
                GetLocalServer();
            }
        }


        /// <summary>
        /// 连接服务器
        /// </summary>
        /// <param name="remoteServerIP">服务器IP</param>
        /// <param name="remoteServerName">服务器名称</param>
        /// <returns></returns>
        public bool ConnectRemoteServer(string remoteServerIP, string remoteServerName)
        {
            try
            {
                KepServer.Connect(remoteServerName, remoteServerIP);
                if (KepServer.ServerState != (int)OPCServerState.OPCRunning) return false;
            }
            catch
            {
                return false;
            }
            return true;
        }
        /// <summary>
        /// 创建组，将本地组和服务器上的组对应
        /// </summary>
        /// <returns></returns>
        public bool CreateGroup()
        {
            try
            {
                KepGroups = KepServer.OPCGroups;//将服务端的组集合复制到本地
                KepGroup = KepGroups.Add("S");//添加一个组
                KepServer.OPCGroups.DefaultGroupIsActive = Convert.ToBoolean(true);//激活组
                KepServer.OPCGroups.DefaultGroupDeadband = Convert.ToInt32(0);// 死区值，设为0时，服务器端该组内任何数据变化都通知组
                KepGroup.UpdateRate = Convert.ToInt32(5000);//服务器向客户程序提交数据变化的刷新速率
                KepGroup.IsActive = Convert.ToBoolean(true);//组的激活状态标志
                KepGroup.IsSubscribed = Convert.ToBoolean(true);//是否订阅数据
                KepItems = KepGroup.OPCItems;//将组里的节点集合复制到本地节点集合
                KepGroup.DataChange += KepGroup_DataChange;
                List<string> ItemIds = new List<string>()
                {
                   ConfigurationManager.AppSettings["Opc_Speed"],
                   ConfigurationManager.AppSettings["Opc_Flux"],
                   ConfigurationManager.AppSettings["Opc_Load"],
                   ConfigurationManager.AppSettings["Opc_Si"],
                   ConfigurationManager.AppSettings["Opc_Al"],
                   ConfigurationManager.AppSettings["Opc_Fe"],
                   ConfigurationManager.AppSettings["Opc_Ca"],
                   ConfigurationManager.AppSettings["Opc_Mg"],
                   ConfigurationManager.AppSettings["Opc_K"],
                   ConfigurationManager.AppSettings["Opc_Na"],
                   ConfigurationManager.AppSettings["Opc_S"],
                   ConfigurationManager.AppSettings["Opc_Cl"],
                };
                for (int i = 0; i < ItemIds.Count; i++)
                {
                    OPCItem myItem = KepGroup.OPCItems.AddItem(ItemIds[i], i);
                }
            }
            catch
            {
                return false;
            }
            return true;
        }
        /// <summary>
        /// 数据订阅方法
        /// </summary>
        /// <param name="TransactionID">处理ID</param>
        /// <param name="NumItems">项个数</param>
        /// <param name="ClientHandles">OPC客户端的句柄</param>
        /// <param name="ItemValues">节点的值</param>
        /// <param name="Qualities">节点的质量</param>
        /// <param name="TimeStamps">时间戳</param>
        private void KepGroup_DataChange(int TransactionID, int NumItems, ref Array ClientHandles, ref Array ItemValues, ref Array Qualities, ref Array TimeStamps)
        {
            for (int i = 1; i <= NumItems; i++)//下标一定要从1开始，NumItems参数是每次事件触发时Group中实际发生数据变化的Item的数量，而不是整个Group里的Items
            {
                list[i - 1] = ItemValues.GetValue(i).ToString();
            }
            list[12] = DateTime.Now.ToString();

            if (list[1] != "0")
            {
                Queue.Enqueue(list);
            }
        }
        /// <summary>
        /// Gama存储至服务器
        /// </summary>
        public void GamaUpload()
        {
            while (opc_connected)
            {
                try
                {
                    if (Queue.Count > 0)
                    {
                        Queue.TryDequeue(out List<string> temp);
                        string sqls = "insert into gama_orig(GAMA_FLUX,GAMA_LOAD,GAMA_SI,GAMA_AL,GAMA_FE,GAMA_CA,GAMA_MG,GAMA_K,GAMA_NA,GAMA_S,GAMA_CL,COMPANY,ADD_TIME)values('" + temp[1] + "','" + temp[2] + "','" + temp[3] + "','" + temp[4] + "','" + temp[5] + "','" + temp[6] + "','" + temp[7] + "','" + temp[8] + "','" + temp[9] + "','" + temp[10] + "','" + temp[11] + "','" + Company + "','" + temp[12] + "')";
                        MySqlHelpers.UpdOrInsOrdel(sqls);
                    }
                    else
                    {
                        Thread.Sleep(5000);
                    }
                }
                catch
                {
                    Gama.sendMessage("存储错误");
                }
            }
        }
    }
}
