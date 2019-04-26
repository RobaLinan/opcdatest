using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OPCAutomation;

namespace opcdatest
{
    class Program
    {
        private static OPCServer server = new OPCServer();
        private static string host;
        private static Array AllOpcServers;
        private static int serverIndex;
        private static string browserFilter = "";
        private static OPCBrowser browser;
        private static string currentPosition = "";
        private static List<string> itemArray;
        private static List<int> clientHandleArray;
        private static OPCGroup group;
        private static bool eventRegisted = false;

        /// <summary>
        /// opc服务所在的主机名/ip/url
        /// </summary>
        /// <returns></returns>
        private static string ChangeHost()
        { 
            Console.Write("请输入计算机名或IP地址, 默认使用localhost: ");
            string newhost = Console.ReadLine();
            newhost = string.IsNullOrWhiteSpace(newhost) ? "localhost" : newhost;
            return newhost;
        }

        /// <summary>
        /// 在host上搜索OPC服务器
        /// </summary>
        /// <param name="hostname"></param>
        private static void SearchServers(string hostname)
        { 
            host = string.IsNullOrWhiteSpace(host) ? "localhost" : host;
            object AllOpcServersArr;
            try
            {
                AllOpcServersArr = server.GetOPCServers(host);
            }
            catch
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("在 {0} 上未发现OPC服务! ", host);
                Console.ResetColor();
                return;
            }
            AllOpcServers = (Array)AllOpcServersArr;
            ListServers();
        }

        /// <summary>
        /// 搜索到的服务器名保存在 AllOpcServers 中
        /// </summary>
        private static void ListServers()
        {
            Console.WriteLine("在 {0} 上找到 {1} 个OPC服务器: ", host, AllOpcServers.Length);
            for (int i = AllOpcServers.GetLowerBound(0); i <= AllOpcServers.Length; i++)
            {
                Console.Write(i.ToString() + ": ");
                Console.WriteLine(AllOpcServers.GetValue(i));
            }

        }

        /// <summary>
        /// 连接OPC服务器
        /// </summary>
        private static void ConnectToServer()
        {
            if (AllOpcServers == null)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("先搜索服务器!");
                Console.ResetColor();
                return;
            }
            ListServers();
            Console.Write("选择一个服务器: ");
            bool indexRight = int.TryParse(Console.ReadLine(), out serverIndex);
            if (indexRight && serverIndex <= AllOpcServers.Length)
            {
                try
                {
                    server.Connect(AllOpcServers.GetValue(serverIndex).ToString(), host);
                }
                catch
                {
                    Console.ForegroundColor = ConsoleColor.Red;
                    Console.WriteLine("连接 {0} 失败!", AllOpcServers.GetValue(serverIndex));
                    Console.ResetColor();
                    return;
                }
                if (server.ServerState == 1)
                {
                    Console.WriteLine("连接 {0} 成功!", server.ServerName);
                }
                else
                {
                    Console.ForegroundColor = ConsoleColor.Red;
                    Console.WriteLine("连接 {0} 失败!", AllOpcServers.GetValue(serverIndex));
                    Console.ResetColor();
                    return;
                }
            }
            else
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("输入错误!");
                Console.ResetColor();
                return;
            }
        }

        /// <summary>
        /// 设置OPCBrowser筛选字符, 前后加 * 作为通配符
        /// </summary>
        /// <returns></returns>
        private static string SetBrowserFilter()
        {
            Console.Write("输入筛选字符: ");
            return "*" + Console.ReadLine() + "*";
        }

        private static void BrowseItems()
        {
            if (server.ServerState != 1)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("先连接到服务器!");
                Console.ResetColor();
                return;
            }
            if(eventRegisted)
                UnRegistDataChangeEvent();
            browser = server.CreateBrowser();
            browser.MoveTo(CurrentPosotion(currentPosition));
            browser.Filter = browserFilter;
            browser.ShowBranches();
            browser.ShowLeafs(true);
            itemArray = new List<string>();
            foreach (var item in browser)
            {
                if (!item.ToString().Contains("Hint"))
                    itemArray.Add(item.ToString());
            }
            Console.WriteLine("找到 {0} 个变量.", itemArray.Count);
        }

        /// <summary>
        /// 打印变量列表
        /// </summary>
        private static void ListBrowserItems()
        {
            if (browser == null)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("先浏览变量!");
                Console.ResetColor();
                return;
            }
            BrowseItems();
            //当前在根节点, 直接列出变量
            if (string.IsNullOrEmpty(currentPosition))
            {
                foreach (var item in itemArray)
                {
                    Console.WriteLine(item);
                }
            }
            //当前在分支节点, 以树形结构列出变量
            else
            {
                Console.Write(currentPosition);
                foreach (var item in itemArray)
                {
                    if (itemArray.IndexOf(item) == 0)
                    {
                        Console.Write("--");
                    }
                    else
                    {
                        foreach(char c in currentPosition)
                            Console.Write(" ");
                        Console.Write("|-");
                    }
                    Console.WriteLine(item.Remove(0, currentPosition.Length + 1));
                }
            }
            Console.WriteLine("找到 {0} 个变量.", itemArray.Count);
        }

        /// <summary>
        /// 将当前节点转为Array以便OPCBrowser.Move使用
        /// </summary>
        /// <param name="branch">节点字符串</param>
        /// <returns>Array</returns>
        private static Array CurrentPosotion(string branch)
        {
            string channel, device, group;
            if (branch.Contains('.'))
            {
                channel = branch.Substring(0, branch.IndexOf('.'));
                branch = branch.Remove(branch.IndexOf(channel), channel.Length + 1);
                if (branch.Contains('.'))
                {
                    device = branch.Substring(0, branch.IndexOf('.'));
                    branch = branch.Remove(branch.IndexOf(device), device.Length + 1);
                }
                else
                {
                    device = branch;
                    branch = branch.Remove(branch.IndexOf(device), device.Length);
                }
                group = branch;
            }
            else
            {
                channel = branch;
                device = "";
                group = "";
            }
            Array branchArray = Array.CreateInstance(typeof(string), new int[] { 3 }, new int[] { 1 });
            branchArray.SetValue(channel, 1);
            branchArray.SetValue(device, 2);
            branchArray.SetValue(group, 3);

            return branchArray;
        }

        /// <summary>
        /// 移动OPCBrowser节点
        /// </summary>
        private static void MoveToBranch()
        {
            if (browser == null)
            { 
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("先浏览变量!");
                Console.ResetColor();
                return;
            }
            Console.Write("输入分枝名: ");
            string branch = Console.ReadLine();
            Array branchArray = Array.CreateInstance(typeof(string), new int[] { 3 }, new int[] { 1 });
            branchArray = CurrentPosotion(branch);
            try
            {
                browser.MoveTo(ref branchArray);
            }
            catch
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("输入错误!");
                Console.ResetColor();
                browser.MoveToRoot();
            }
            Console.WriteLine("移动至 " + browser.CurrentPosition);
            currentPosition = browser.CurrentPosition;
        }

        /// <summary>
        /// 注册OPC变量
        /// </summary>
        private static void RegistItems()
        {
            int count = 0;
            if (itemArray == null)
            { 
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("先浏览变量!");
                Console.ResetColor();
                return;
            }
            //if (server.OPCGroups.Count > 0)
            //    server.OPCGroups.RemoveAll();
            //此处用RemoveAll方法不行, 只好重新new一个OPCServer再Connect
            server = new OPCServer();
            server.Connect(AllOpcServers.GetValue(serverIndex).ToString(), host);
            group = server.OPCGroups.Add("Group1");
            //clientHandleArray的作用是在变量变化事件中由clientHandle对应变量名
            clientHandleArray = new List<int>();
            foreach (var item in itemArray)
            {
                group.OPCItems.AddItem(item, count);
                clientHandleArray.Add(count);
                count++;
            }
            //使用订阅方式, 变量发生变化会产生事件
            group.IsSubscribed = true;
            group.IsActive = true;
            group.UpdateRate = 100;
            Console.WriteLine(count.ToString() + " 个变量已注册");
        }

        /// <summary>
        /// 注册变量变化事件
        /// </summary>
        private static void RegistDataChangeEvent()
        {
            if (group == null)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("先注册变量!");
                Console.ResetColor();
                return;
            }
            group.DataChange += ObjOPCGroup_DataChange;
            Console.WriteLine("事件已注册.");
            eventRegisted = true;
        }

        /// <summary>
        /// 注销事件
        /// </summary>
        private static void UnRegistDataChangeEvent()
        { 
            if (group == null)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("先注册变量!");
                Console.ResetColor();
                return;
            }
            if (!eventRegisted)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("事件未注册.");
                Console.ResetColor();
                return;
            }
            group.DataChange -= ObjOPCGroup_DataChange;
            Console.WriteLine("事件已注销.");
        }

        /// <summary>
        /// 变量变化事件实体
        /// </summary>
        /// <param name="TransactionID"></param>
        /// <param name="NumItems">变化的变量个数</param>
        /// <param name="ClientHandles">变化的变量ClientHandle数组</param>
        /// <param name="ItemValues">变化的变量值数组</param>
        /// <param name="Qualities">通信质量数组</param>
        /// <param name="TimeStamps">时间戳数组</param>
        private static void ObjOPCGroup_DataChange(int TransactionID, int NumItems, ref Array ClientHandles, ref Array ItemValues, ref Array Qualities, ref Array TimeStamps)
        {
            Console.WriteLine();
            Console.WriteLine("{0} 个变量变化: ", NumItems);
            for (int i = 1; i <= NumItems; i++)
            {
                int handle = int.Parse(ClientHandles.GetValue(i).ToString());
                int index = clientHandleArray.IndexOf(handle);
                string itemName = itemArray[index];
                string itemValue = ItemValues.GetValue(i).ToString();
                Console.WriteLine("{0} 当前值: {1}", itemName, itemValue);
            }
        }

        /// <summary>
        /// 按名称写入一个OPC变量的值
        /// </summary>
        private static void WriteItem()
        {
            Console.Write("输入要写入的变量名: ");
            string itemName = string.IsNullOrEmpty(currentPosition) ? Console.ReadLine() : currentPosition + "." + Console.ReadLine();
            if (!itemArray.Contains(itemName))
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("未注册变量 {0} !", itemName);
                Console.ResetColor();
                return;
            }
            Console.Write("输入要写入的变量值: ");
            string itemValue = Console.ReadLine();
            try
            {
                group.OPCItems.Item(itemName).Write(itemValue);
            }
            catch (Exception e)
            {
                Console.WriteLine(e.ToString());
            }
        }

        static void Main(string[] args)
        {
            string cmd;
            while (true)
            {
                Console.ForegroundColor = ConsoleColor.Green;
                Console.Write("输入命令 (? for help): ");
                cmd = Console.ReadLine();
                Console.ResetColor();

                switch (cmd.ToLower())
                {
                    case "host":
                    case "h":
                        host = ChangeHost();
                        break;
                    case "search":
                    case "s":
                        SearchServers(host);
                        break;
                    case "c":
                    case "connect":
                        ConnectToServer();
                        break;
                    case "f":
                    case "filter":
                        browserFilter = SetBrowserFilter();
                        break;
                    case "b":
                    case "browse":
                        BrowseItems();
                        break;
                    case "l":
                    case "list":
                        ListBrowserItems();
                        break;
                    case "m":
                    case "move":
                        MoveToBranch();
                        break;
                    case "r":
                    case "regist":
                        RegistItems();
                        break;
                    case "e":
                    case "event":
                        RegistDataChangeEvent();
                        break;
                    case "u":
                    case "unregist":
                        UnRegistDataChangeEvent();
                        break;
                    case "w":
                    case "write":
                        WriteItem();
                        break;
                    case "x":
                    case "exit":
                        return;
                    default:
                        ShowHelp();
                        break;
                }

            }
        }

        static void ShowHelp()
        {
            Console.ForegroundColor = ConsoleColor.Yellow;
            Console.WriteLine("  可以使用[]内的缩写 ");
            Console.WriteLine("     [h]ost          OPC服务所在的主机名/ip地址/URL ");
            Console.WriteLine("     [s]earch        搜索OPC服务器 ");
            Console.WriteLine("     [c]onnect       连接到一个OPC服务器 ");
            Console.WriteLine("     [f]ilter        设置变量筛选 ");
            Console.WriteLine("     [b]rowse        浏览当前分支中的变量 ");
            Console.WriteLine("     [m]ove          移动到一个分支 ");
            Console.WriteLine("     [l]ist          显示变量列表 ");
            Console.WriteLine("     [r]egist        在服务器上注册变量 ");
            Console.WriteLine("     [e]vent         注册变量变化事件 ");
            Console.WriteLine("     [u]negist       注销变量变化事件 ");
            Console.WriteLine("     [w]rite         写变量值 ");

            Console.WriteLine("     e[x]it          退出 ");
            Console.WriteLine("     其它            菜单 ");
            Console.ResetColor();
        }
    }
}
