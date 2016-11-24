using System;
using System.Data;
using System.Configuration;
using System.Collections;
using System.Web;
using System.Web.Security;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Web.UI.WebControls.WebParts;
using System.Web.UI.HtmlControls;
using System.Net;
using System.Net.Sockets;
using System.Net.NetworkInformation;
using System.Text;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Management;
using System.Threading;
using System.Drawing;

public partial class _Default : System.Web.UI.Page
{
    [DllImport("Iphlpapi.dll")]
    static extern int SendARP(Int32 DestIP, Int32 SrcIP, ref Int64 MacAddr, ref Int32 PhyAddrLen);

    [DllImport("Ws2_32.dll")]
    static extern Int32 inet_addr(string ipaddr);

    public delegate void UpdateList(string sIP, string sHostName, string sDomain, string sUsername, string sMac, int iTtl, int lTime, string sSubnet, string sGateway, string sDns1, string sDns2, string sDns3, string sWins1, string sWins2, string sTime);

    static int scannedCount, runningThreadCount, online, compSum, NetTimeout = 100;
    static string WmiScan, DnsScan;
    static TimeSpan WMITimeout = new TimeSpan(0, 0, 0, 30, 0);

    public DateTime dtAuto;
    public string[] compManager = new string[256];
    public ArrayList CompterIPC = new ArrayList();

    public int maxIP4 = 254, maxThread = 16, scanSum, managerCount, TableID = 1;
    public int taskA, taskB, taskC, taskD, taskE, taskF, taskG;
    public string[,] AutoTaskA = new string[256, 3], AutoTaskB = new string[256, 3], AutoTaskC = new string[256, 3], AutoTaskD = new string[256, 3],
        AutoTaskE = new string[256, 3], AutoTaskF = new string[256, 3], AutoTaskG = new string[256, 3];


    public DataTable dt1;
    public ArrayList ChkArraylist;

    protected void Page_Load(object sender, EventArgs e)
    {
        try
        {
            this.Button2.Visible = false;
            this.Button3.Visible = false;
            if (!Page.IsPostBack)
            {
                this.TextBox1.Text = Request.UserHostAddress.ToString();
                /*try
                {
                    ManagementObjectSearcher vManagementObjectSearcher = new ManagementObjectSearcher(@"root\WMI", @"select * from MSAcpi_ThermalZoneTemperature");
                    foreach (ManagementObject vManagementObject in vManagementObjectSearcher.Get())
                    {   //-2732再除以10
                        double temperature = double.Parse(vManagementObject.Properties["CurrentTemperature"].Value.ToString());
                        temperature = (temperature - 2732) / 10;
                        this.Label2.Text = ("CPU温度为：" + temperature.ToString() + " ℃");
                    }
                }
                catch { }*/
            }
            else
            {
                CreateDT();
                ChkArraylist = new ArrayList();
                if (CheckBox1.Checked == true) { DnsScan = "是"; } else { DnsScan = "否"; }
                if (CheckBox2.Checked == true) { WmiScan = "是"; } else { WmiScan = "否"; }
            }
        }
        catch (Exception ex) { Response.Write(ex.Message); }
    }

    //扫描
    public class LanPing
    {
        public UpdateList ul;
        public string tb_admin, tb_password;
        public string ip, hostname, domain, username, mac;
        public int ttl;
        public int roundtriptime;

        public string hostname1 = "", domain1 = "", username1 = "";
        public string subnet1 = "";
        public string gateway1 = "";
        public string dns1 = "", dns2 = "", dns3 = "";
        public string wins1 = "", wins2 = "";
        public string macaddress1 = "";

        public int bar;

        public bool boolLanIP;

        public void Scan()
        {
            IPAddress myIP = IPAddress.Parse(ip);
            string tempIP = "";

            Ping pingSender = new Ping();
            PingOptions options0 = new PingOptions();

            options0.DontFragment = true;

            string data = "#012345678901234567890123456789*";
            byte[] buffer = Encoding.ASCII.GetBytes(data);
            int timeout = NetTimeout;
            PingReply reply = pingSender.Send(myIP, timeout, buffer, options0);

            if (reply.Status == IPStatus.Success)
            {
                try
                {
                    roundtriptime = (int)reply.RoundtripTime;
                    ttl = reply.Options.Ttl;
                    string[] HostInfo = Nbtstat(ip);
                    hostname = HostInfo[0].Trim();
                    domain = HostInfo[1].Trim();
                    username = HostInfo[2].Trim();
                    mac = HostInfo[3].Trim();
                }
                catch
                { }

                if (string.IsNullOrEmpty(hostname) && DnsScan == "是")
                {
                    try
                    {
                        hostname = Dns.GetHostEntry(myIP).HostName.ToString();
                        if (hostname == ip) { hostname = ""; }
                    }
                    catch { hostname = ""; }
                }

                if (boolLanIP == true && string.IsNullOrEmpty(mac))
                {
                    try { mac = GetMacAddress(ip); }
                    catch { }
                }

                if (mac == "00-00-00-00-00-00") { mac = ""; }

                #region WMI扫描
                if (!string.IsNullOrEmpty(hostname) && WmiScan == "是")
                {
                    ConnectionOptions options = new ConnectionOptions();
                    options.Timeout = WMITimeout;

                    IPAddress[] hostip = Dns.GetHostEntry(Dns.GetHostName()).AddressList;
                    for (int i = 0; i < hostip.Length; i++)
                    {
                        if (hostip[i].ToString() == ip)
                        {
                            tempIP = ip;
                            ip = ".";
                        }
                    }

                    if (ip != ".")
                    {
                        try
                        {
                            options.Username = tb_admin;
                            options.Password = tb_password;
                        }
                        catch { }
                    }

                    ManagementScope scope = new ManagementScope("\\\\" + myIP + "\\root\\cimv2", options);

                    if (ip != "." && !string.IsNullOrEmpty(tb_admin))
                    {
                        try { scope.Connect(); }
                        catch { }
                    }

                    if (scope.IsConnected || ip == ".")
                    {
                        ObjectQuery oq = new ObjectQuery("SELECT * FROM Win32_ComputerSystem");
                        ManagementObjectSearcher query = new ManagementObjectSearcher(scope, oq);
                        ManagementObjectCollection queryCollection = query.Get();

                        foreach (ManagementObject mo in queryCollection)
                        {
                            //机器名
                            try { hostname1 = mo["Name"].ToString(); }
                            catch { }
                            //工作组（域）
                            try { domain1 = mo["Domain"].ToString(); }
                            catch { }
                            //用户名
                            try { username1 = mo["UserName"].ToString(); }
                            catch { }
                        }

                        oq = new ObjectQuery("SELECT * FROM Win32_NetworkAdapterConfiguration Where IPEnabled = 'true'");
                        query = new ManagementObjectSearcher(scope, oq);
                        queryCollection = query.Get();

                        foreach (ManagementObject mo in queryCollection)
                        {
                            //子网掩码
                            try
                            {
                                string[] subnets = (string[])mo["IPSubnet"];
                                subnet1 = subnets[0];
                            }
                            catch { }
                            //默认网关
                            try
                            {
                                string[] gateways = (string[])mo["DefaultIPGateway"];
                                gateway1 = gateways[0];
                            }
                            catch { }
                            //dns解析服务器
                            try
                            {
                                string[] dnses = (string[])mo["DNSServerSearchOrder"];
                                dns1 = dnses[0];
                                dns2 = dnses[1];
                                dns3 = dnses[2];
                            }
                            catch { }
                            //Wins服务器1
                            try { wins1 = mo["WINSPrimaryServer"].ToString(); }
                            catch { }
                            //Wins服务器2
                            try { wins2 = mo["WINSSecondaryServer"].ToString(); }
                            catch { }
                            //Mac地址
                            try { macaddress1 = mo["MACAddress"].ToString().Replace(':', '-'); }
                            catch { }
                        }
                    }
                }
                #endregion

                if (ip == ".") { ip = tempIP; }

                if (!string.IsNullOrEmpty(hostname1)) { hostname = hostname1; }
                if (!string.IsNullOrEmpty(domain1)) { domain = domain1; }
                if (!string.IsNullOrEmpty(username1)) { username = username1; }
                if (!string.IsNullOrEmpty(macaddress1)) { mac = macaddress1; }

                if (ul != null)
                {
                    try
                    {
                        ul(ip, hostname, domain, username, mac, ttl, roundtriptime, subnet1, gateway1, dns1, dns2, dns3, wins1, wins2, DateTime.Now.ToString());
                        online++;
                        if (!string.IsNullOrEmpty(hostname)) { compSum++; }
                    }
                    catch { }
                }
            }

            scannedCount++;
            runningThreadCount--;
        }
    }

    //委托更新主窗口数据表格
    void UpdateMyList(string sIP, string sHostName, string sDomain, string sUsername, string sMac, int iTtl, int lTime, string sSubnet, string sGateway, string sDns1, string sDns2, string sDns3, string sWins1, string sWins2, string sTime)
    {
        lock (this.GridView1)
        {
            if (!string.IsNullOrEmpty(sIP))
            {
                dt1.Rows.Add(new object[] { TableID, sIP, sHostName, sDomain, sUsername, sMac, iTtl, lTime, sSubnet, sGateway, sDns1, sDns2, sDns3, sWins1, sWins2, sTime, false });
                TableID++;
            }
        }
    }

    bool IsLanIP(string LanIP)
    {
        bool BoolLan = false;
        try
        {
            IPAddress[] LocalIPs = Dns.GetHostEntry(Dns.GetHostName()).AddressList;

            foreach (IPAddress LocalIP in LocalIPs)
            {
                if (LocalIP.ToString().IndexOf(':') == -1)
                {
                    if (LanIP.Substring(0, LanIP.LastIndexOf('.')) == LocalIP.ToString().Substring(0, LocalIP.ToString().LastIndexOf('.')))
                    {
                        BoolLan = true;
                    }
                }

            }
        }
        catch (Exception ex) { this.Label2.Text = "是否局域网IP地址：" + ex.Message.ToString(); }

        return BoolLan;
    }

    //网段扫描
    protected void Button1_Click(object sender, EventArgs e)
    {
        try
        {
            online = 0;
            compSum = 0;
            scannedCount = 0;
            runningThreadCount = 0;

            int Min = 1, Max = 254;

            scanSum = Max - Min + 1;

            int _ThreadNum = Max - Min + 1;
            Thread[] mythread = new Thread[_ThreadNum];
            string[] TextIP = this.TextBox1.Text.Split('.');
            for (int i = Min; i <= Max; i++)
            {
                int k = Max - i;
                LanPing HostPing = new LanPing();
                HostPing.ip = TextIP[0] + "." + TextIP[1] + "." + TextIP[2] + "." + i.ToString();
                HostPing.boolLanIP = IsLanIP(HostPing.ip);
                HostPing.ul = new UpdateList(UpdateMyList);
                HostPing.tb_admin = TextBox2.Text.Trim();
                HostPing.tb_password = TextBox3.Text.Trim();
                mythread[k] = new Thread(new ThreadStart(HostPing.Scan));
                mythread[k].Start();

                runningThreadCount++;
                Thread.Sleep(10);
                //循环，直到某个线程工作完毕才启动另一新线程，也可以叫做推拉窗技术
                while (runningThreadCount >= maxThread) ;
            }

            bool ViewUpdate = true;
            while (true)
            {
                ViewUpdate = true;
                for (int n = 0; n < Max - Min + 1; n++)
                {
                    if (mythread[n].IsAlive == true)
                    {
                        ViewUpdate = false;
                        Thread.Sleep(100);
                        break;
                    }
                }
                if (ViewUpdate == true)
                {
                    break;
                }
            }
            dt1.AcceptChanges();
            this.GridView1.DataSource = dt1.DefaultView;
            this.GridView1.DataBind();
            if (this.GridView1.Rows.Count > 0)
            {
                this.Button2.Visible = true;
                this.Button3.Visible = true;
            }
        }
        catch { }
    }


    public static string[] Nbtstat(string macIp2)//C#版 NBTSTAT 版权所有 Mossan 2006
    {
        string[] StrNbtstat = new string[4];

        byte[] ByteSend = new byte[50] { 0x0, 0x00, 0x0, 0x10, 0x0, 0x1, 0x0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x20, 0x43, 0x4b, 0x41, 0x41, 0x41, 0x41, 0x41, 0x41, 0x41, 0x41, 0x41, 0x41, 0x41, 0x41, 0x41, 0x41, 0x41, 0x41, 0x41, 0x41, 0x41, 0x41, 0x41, 0x41, 0x41, 0x41, 0x41, 0x41, 0x41, 0x41, 0x41, 0x41, 0x0, 0x0, 0x21, 0x0, 0x1 };
        byte[] ByteBuf = new byte[500];
        byte[,] ByteRecv = new byte[18, 28];
        string strTemp = "", hostname2 = "", domain2 = "", username2 = "", mac2 = "";
        int iRecv, macline = 0, usernum = 0;
        string[] domainuser = new string[2];
        domainuser[0] = "";
        domainuser[1] = "";

        try
        {
            IPEndPoint ipepSend = new IPEndPoint(IPAddress.Any, 0);
            EndPoint epRemote = (EndPoint)ipepSend;

            IPEndPoint ipepNbt = new IPEndPoint(IPAddress.Parse(macIp2), 137);

            Socket sockServer = new Socket(AddressFamily.InterNetwork, SocketType.Dgram, ProtocolType.Udp);

            sockServer.SetSocketOption(SocketOptionLevel.Socket, SocketOptionName.ReceiveTimeout, NetTimeout);
            sockServer.SendTo(ByteSend, ByteSend.Length, SocketFlags.None, ipepNbt);
            iRecv = sockServer.ReceiveFrom(ByteBuf, ref epRemote);
            sockServer.Close();

            if (iRecv > 0)
            {
                ByteRecv = new byte[18, (iRecv - 56) % 18];

                for (int k = 0; k < (iRecv - 56) % 18; k++)
                {
                    for (int j = 0; j < 18; j++)
                    {
                        ByteRecv[j, k] = ByteBuf[57 + 18 * k + j];
                    }
                }

                for (int k = 0; k < (iRecv - 56) % 18; k++)
                {
                    if (System.Convert.ToString(ByteRecv[15, k], 16) == "0" && (System.Convert.ToString(ByteRecv[16, k], 16) == "4" || System.Convert.ToString(ByteRecv[16, k], 16) == "44"))
                    {
                        strTemp = "";
                        for (int j = 0; j < 15; j++)
                        {
                            strTemp += System.Convert.ToChar(ByteRecv[j, k]).ToString();
                        }
                        hostname2 = strTemp;
                    }

                    if (System.Convert.ToString(ByteRecv[15, k], 16) == "0" && (System.Convert.ToString(ByteRecv[16, k], 16) == "84" || System.Convert.ToString(ByteRecv[16, k], 16).ToUpper() == "C4"))
                    {
                        strTemp = "";
                        for (int j = 0; j < 15; j++)
                        {
                            strTemp += System.Convert.ToChar(ByteRecv[j, k]).ToString();
                        }
                        domain2 = strTemp;
                    }

                    if (System.Convert.ToString(ByteRecv[15, k], 16) == "3" && (System.Convert.ToString(ByteRecv[16, k], 16) == "4" || System.Convert.ToString(ByteRecv[16, k], 16) == "44"))
                    {
                        strTemp = "";
                        for (int j = 0; j < 15; j++)
                        {
                            strTemp += System.Convert.ToChar(ByteRecv[j, k]).ToString();
                        }
                        domainuser[usernum] = strTemp;
                        usernum++;
                    }

                    if (System.Convert.ToString(ByteRecv[15, k], 16) == "0" && System.Convert.ToString(ByteRecv[16, k], 16) == "0" && System.Convert.ToString(ByteRecv[17, k], 16) == "0")
                    {
                        macline = k;

                        for (int i = 0; i < 6; i++)
                        {
                            if (i < 5)
                            {
                                mac2 += System.Convert.ToString(ByteRecv[i, macline], 16).PadLeft(2, '0').ToUpper() + "-";
                            }
                            if (i == 5)
                            {
                                mac2 += System.Convert.ToString(ByteRecv[i, macline], 16).PadLeft(2, '0').ToUpper();
                            }
                        }
                        k = (iRecv - 56) % 18;
                    }
                }
            }
        }
        catch //(SocketException se)
        {
            //MessageBox.Show(se.Message); 
        }
        username2 = domainuser[1];
        if (string.IsNullOrEmpty(domainuser[1])) { username2 = domainuser[0]; }

        StrNbtstat = new string[] { hostname2, domain2, username2, mac2 };
        return StrNbtstat;
    }

    //SendArp获取Mac地址
    public static string GetMacAddress(string macip)
    {
        StringBuilder strReturn = new StringBuilder();
        try
        {
            Int32 remote = inet_addr(macip);

            Int64 macinfo = new Int64();
            Int32 length = 6;
            SendARP(remote, 0, ref macinfo, ref length);

            string temp = System.Convert.ToString(macinfo, 16).PadLeft(12, '0').ToUpper();

            int x = 12;
            for (int i = 0; i < 6; i++)
            {
                if (i == 5) { strReturn.Append(temp.Substring(x - 2, 2)); }
                else { strReturn.Append(temp.Substring(x - 2, 2) + "-"); }
                x -= 2;
            }

            return strReturn.ToString();
        }
        catch
        {
            return strReturn.ToString();
        }
    }

    //GridView1排序
    protected void GridView1_Sorting(object sender, GridViewSortEventArgs e)
    {
        try
        {
            string SortDirect1 = "";
            if (e.SortDirection == SortDirection.Ascending)
            {
                SortDirect1 = " Asc";
            }
            else
            {
                SortDirect1 = " Desc";
            }

            BackupChkBox();
            GetDataFromGrid();
            DataTable dt2 = TableSort(dt1, e.SortExpression + SortDirect1);
            this.GridView1.DataSource = dt2.DefaultView;
            this.GridView1.DataBind();
            RestoreChkBox(dt2.Rows.Count);
        }
        catch (Exception ex)
        {
            this.Label2.Text = "GridView1排序：" + ex.Message.ToString();
        }

        if (this.GridView1.Rows.Count > 0)
        {
            this.Button2.Visible = true;
            this.Button3.Visible = true;
        }

    }

    // 对DataTable中的数据进行排序
    private DataTable TableSort(DataTable dt, string sort)
    {
        DataRow[] resultRows = dt.Select(null, sort);
        DataTable tableTemp = dt.Copy();
        tableTemp.Rows.Clear();
        foreach (DataRow dr in resultRows)
        {
            object[] objItem = dr.ItemArray;
            DataRow row = tableTemp.NewRow();
            row.ItemArray = objItem;
            tableTemp.Rows.Add(row);
        }
        return tableTemp;
    }

    //创建dt1
    public void CreateDT()
    {
        dt1 = new DataTable("Table1");

        DataColumn PrimaryKey1 = new DataColumn();
        PrimaryKey1.DataType = System.Type.GetType("System.Int32");
        PrimaryKey1.ColumnName = "序号";
        dt1.Columns.Add(PrimaryKey1);

        DataColumn[] keys = new DataColumn[1];
        keys[0] = PrimaryKey1;
        dt1.PrimaryKey = keys;

        dt1.Columns.Add("IP地址");
        dt1.Columns.Add("计算机名");
        dt1.Columns.Add("工作组(域)");
        dt1.Columns.Add("登录名");
        dt1.Columns.Add("MAC地址");

        DataColumn ColumnTTL = new DataColumn();
        ColumnTTL.DataType = System.Type.GetType("System.Int32");
        ColumnTTL.ColumnName = "TTL";
        dt1.Columns.Add(ColumnTTL);

        DataColumn ColumnRound = new DataColumn();
        ColumnRound.DataType = System.Type.GetType("System.Int32");
        ColumnRound.ColumnName = "回传";
        dt1.Columns.Add(ColumnRound);

        dt1.Columns.Add("子网掩码");
        dt1.Columns.Add("默认网关");
        dt1.Columns.Add("DNS服务器1");
        dt1.Columns.Add("DNS服务器2");
        dt1.Columns.Add("DNS服务器3");
        dt1.Columns.Add("WINS服务器1");
        dt1.Columns.Add("WINS服务器2");
        dt1.Columns.Add("日期时间");
        dt1.Columns.Add("选择");
    }

    //删除行
    protected void GridView1_RowDeleting(object sender, GridViewDeleteEventArgs e)
    {
        try
        {
            BackupChkBox();
            GetDataFromGrid();
            dt1.Rows.RemoveAt(e.RowIndex);
            this.GridView1.DataSource = dt1.DefaultView;
            this.GridView1.DataBind();
            RestoreChkBox(dt1.Rows.Count);
        }
        catch (Exception ex)
        {
            this.Label2.Text = "删除行：" + ex.Message.ToString();
        }
        if (this.GridView1.Rows.Count > 0)
        {
            this.Button2.Visible = true;
            this.Button3.Visible = true;
        }
    }

    //选择所有
    protected void chkSelectAll_CheckedChanged(object sender, EventArgs e)
    {
        //遍历GridView行获取CheckBox属性
        for (int i = 0; i < this.GridView1.Rows.Count; i++)
        {
            ((CheckBox)GridView1.Rows[i].FindControl("chkSelect")).Checked = ((CheckBox)GridView1.HeaderRow.FindControl("chkSelectAll")).Checked;
        }

        if (this.GridView1.Rows.Count > 0)
        {
            this.Button2.Visible = true;
            this.Button3.Visible = true;
        }

    }

    //选择单个
    protected void chkSelect_CheckedChanged(object sender, EventArgs e)
    {
        int ChkCount = 0;
        //遍历GridView行获取CheckBox属性
        for (int i = 0; i < this.GridView1.Rows.Count; i++)
        {
            if (((CheckBox)GridView1.Rows[i].FindControl("chkSelect")).Checked == false)
            {
                if (((CheckBox)GridView1.HeaderRow.FindControl("chkSelectAll")).Checked == true)
                {
                    ((CheckBox)GridView1.HeaderRow.FindControl("chkSelectAll")).Checked = false;
                }
            }
            else { ChkCount++; }
        }

        if (ChkCount == this.GridView1.Rows.Count)
        {
            ((CheckBox)GridView1.HeaderRow.FindControl("chkSelectAll")).Checked = true;
        }

        if (this.GridView1.Rows.Count > 0)
        {
            this.Button2.Visible = true;
            this.Button3.Visible = true;
        }

    }

    //删除所选
    protected void Button2_Click(object sender, EventArgs e)
    {
        try
        {
            GetDataFromGrid();

            ArrayList Array1 = new ArrayList();

            for (int i = 0; i < this.GridView1.Rows.Count; i++)
            {
                if (((CheckBox)GridView1.Rows[i].FindControl("chkSelect")).Checked == true)
                {
                    Array1.Add(this.GridView1.Rows[i].RowIndex);
                }
            }

            Array1.Sort();
            Array1.Reverse();

            for (int i = 0; i < Array1.Count; i++)
            {
                dt1.Rows.RemoveAt(Convert.ToInt32(Array1[i]));
            }

            this.GridView1.DataSource = dt1.DefaultView;
            this.GridView1.DataBind();
        }
        catch (Exception ex)
        {
            this.Label2.Text = "删除所选：" + ex.Message.ToString();
        }

        if (this.GridView1.Rows.Count > 0)
        {
            this.Button2.Visible = true;
            this.Button3.Visible = true;
        }

    }

    //从GridView获得dt1数据
    protected DataTable GetDataFromGrid()
    {
        int ColumnsCount = this.GridView1.Columns.Count;
        int RowsCount = this.GridView1.Rows.Count;

        string[,] SortGV = new string[ColumnsCount, RowsCount];

        dt1.Clear();
        for (int m = 0; m < RowsCount; m++)
        {
            for (int i = 0; i < ColumnsCount; i++)
            {
                SortGV[i, m] = this.GridView1.Rows[m].Cells[i].Text;
                if (SortGV[i, m] == "&nbsp;")
                {
                    SortGV[i, m] = "";
                }
            }
            dt1.Rows.Add(new object[] { SortGV[0, m], SortGV[1, m], SortGV[2, m], SortGV[3, m], SortGV[4, m], SortGV[5, m], SortGV[6, m], SortGV[7, m], SortGV[8, m], SortGV[9, m], SortGV[10, m], SortGV[11, m], SortGV[12, m], SortGV[13, m], SortGV[14, m], SortGV[15, m], ((CheckBox)GridView1.Rows[m].FindControl("chkSelect")).Checked });
        }
        dt1.AcceptChanges();
        return dt1;
    }

    //保存CheckBox状态
    private void BackupChkBox()
    {
        ChkArraylist.Add(((CheckBox)GridView1.HeaderRow.FindControl("chkSelectAll")).Checked);
        for (int i = 0; i < this.GridView1.Rows.Count; i++)
        {
            if (((CheckBox)GridView1.Rows[i].FindControl("chkSelect")).Checked == true)
            {
                ChkArraylist.Add(this.GridView1.Rows[i].Cells[0].Text);
            }
        }
    }

    //恢复CheckBox状态
    private void RestoreChkBox(int RowsCount)
    {
        ((CheckBox)GridView1.HeaderRow.FindControl("chkSelectAll")).Checked = (bool)ChkArraylist[0];
        for (int n = 1; n < ChkArraylist.Count; n++)
        {
            for (int i = 0; i < RowsCount; i++)
            {
                if (this.GridView1.Rows[i].Cells[0].Text == ChkArraylist[n].ToString())
                {
                    ((CheckBox)GridView1.Rows[i].FindControl("chkSelect")).Checked = true;
                }
            }
        }
    }

    public override void VerifyRenderingInServerForm(Control control)
    { }

    protected void Button3_Click(object sender, EventArgs e)
    {
        if (this.GridView1.Rows.Count > 0)
        {
            string LanScan = this.GridView1.Rows[0].Cells[1].Text;
            LanScan = LanScan.Substring(0, LanScan.LastIndexOf('.')).Replace('.', '_');
            string filename = LanScan + "_" + DateTime.Now.Year.ToString() + DateTime.Now.Month.ToString().PadLeft(2, '0') + DateTime.Now.Day.ToString().PadLeft(2, '0')
                + DateTime.Now.Hour.ToString().PadLeft(2, '0') + DateTime.Now.Minute.ToString().PadLeft(2, '0') + DateTime.Now.Second.ToString().PadLeft(2, '0') + ".xls";
            this.GridView1.Columns[16].Visible = false;
            this.GridView1.Columns[17].Visible = false;
            Response.Clear();
            Response.Buffer = true;
            Response.Charset = "GB2312";
            Response.AppendHeader("Content-Disposition", "attachment;filename=" + filename);
            // 如果设置为 GetEncoding("GB2312");导出的文件将会出现乱码！！！
            //Response.ContentEncoding = System.Text.Encoding.UTF7;
            Response.ContentEncoding = System.Text.Encoding.GetEncoding("GB2312");
            Response.ContentType = "application/ms-excel";//设置输出文件类型为excel文件。 
            System.IO.StringWriter oStringWriter = new System.IO.StringWriter();
            System.Web.UI.HtmlTextWriter oHtmlTextWriter = new System.Web.UI.HtmlTextWriter(oStringWriter);
            this.GridView1.RenderControl(oHtmlTextWriter);
            Response.Output.Write(oStringWriter.ToString());
            Response.Flush();
            Response.End();
            this.GridView1.Columns[16].Visible = true;
            this.GridView1.Columns[17].Visible = true;
        }
    }

    protected void GridView1_RowDataBound(object sender, GridViewRowEventArgs e)
    {
        if (e.Row.RowType == DataControlRowType.DataRow)
        {
            e.Row.Cells[1].Attributes.Add("style", "vnd.ms-excel.numberformat:@;");
            //加入鼠标滑过的高亮效果
            //当鼠标放上去的时候 先保存当前行的背景颜色 并给附一颜色
            //e.Row.Attributes.Add("onmouseover", "currentcolor=this.style.backgroundColor;this.style.backgroundColor='yellow',this.style.fontWeight='';");
            //当鼠标离开的时候 将背景颜色还原的以前的颜色
            //e.Row.Attributes.Add("onmouseout", "this.style.backgroundColor=currentcolor,this.style.fontWeight='';");
            //单击行改变行背景颜色
            e.Row.Attributes.Add("onclick", "currentcolor=this.style.backgroundColor;this.style.backgroundColor='yellow'; this.style.color='buttontext';this.style.cursor='default';");
            //双击行还原行背景颜色
            e.Row.Attributes.Add("onDblClick", "this.style.backgroundColor='',this.style.fontWeight='';");
        }
    }

}
