﻿using System;
using System.Collections;//ArrayList
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.IO;
//Process需要引入System.Diagnostics 和 System.Management
using System.Diagnostics;
using System.Management;
using System.Data.OleDb;
using System.Windows.Forms;

//DataSet
using System.Data;
using System.Threading;

//HL7的引用
using System.Xml;
using System.Net;
using System.Net.Sockets;

//定时器
using System.Timers;

//使用系统电源模式API函数
using System.Runtime.InteropServices;
using System.Reflection;
using Excel = Microsoft.Office.Interop.Excel;


namespace Seekya
{
    /// <summary>
    /// MainWindow.xaml 的交互逻辑
    /// </summary>
    public partial class MainWindow : Window
    {
        private bool winSizeState=true;//窗口状态正常大小，最大化时，赋值为False
        private String serialNum = "";//存储条形码号
        private TcpListener tcpLister = null;

        //远程连接的定时器
        System.Timers.Timer aTimer = null;

        //定义API函数
        [DllImport("kernel32.dll")]
        static extern uint SetThreadExecutionState(uint esFlags);
        const uint ES_SYSTEM_REQUIRED = 0X00000001;
        const uint ES_DISPLAY_REQUIRED = 0X00000002;
        const uint ES_CONTINUOUS = 0X80000000;

        //存储当天数据库的最新行数
        int dbRows = 0;

        //定义一个空的质控窗口类
        QC myQC = null;

        //定义一个空的配置窗口类
        configForm myCfg = null;

        //定义一个空的数据库管理窗口类
        dbManager myDb = null;

        //存储血红蛋白浓度
        public string rbcon = null;

        //子窗口打开标志位
        public bool qcOpen = false;
        public bool cfgOpen = false;
        public bool dbOpen = false;

        public static bool qcOpend = false;

        public bool softwareOperate = false;

        public bool sn = false;

        //定义打印模式标志
        public bool prtmd = true;

        public bool hbmark = true;

        //记录患者信息中需要打印的项目
        string[] values = new string[12];
        //记录患者信息项目
        string[] propts = new string[12];

        //需要监控hb的患者登记号
        public string[] idlist = new string[20];
        //信息来源医院后台的患者数目
        int num = -1;
        //使用医院后台数据的标志
        bool websign = false;
        //判断是否记录临时红细胞寿命的标志，在调用医院后台数据的时候为true,测量完毕为false
        bool wsn = false;
        //使用调用医院后台功能时以血红蛋白为100记录临时红细胞寿命
        int tmpRBC = 0;
        //临时红细胞寿命组
         int[] tmpRBClist = new int[20];
        //临时红细胞寿命对应的检测记录的时间，发给医院后台定位调用血红蛋白对应的那条结果
        string[] timelist = new string[20];
        public MainWindow()
        {
            Process instance = RunningInstance();
            if (instance != null)
            {
                if (instance.MainWindowHandle.ToInt32() == 0)
                {
                    System.Windows.MessageBox.Show("程序已打开并托盘化");
                    return;
                }
                HandleRunningInstance(instance);
            }
            return;

            InitializeComponent();

        }
        #region 确保程序只运行一个程序
        private static Process RunningInstance()
        {
            Process current = Process.GetCurrentProcess();
            Process[] processes = Process.GetProcessesByName(current.ProcessName);
            foreach (Process process in processes)
            {
                if (process.Id != current.Id)
                {
                    if (Assembly.GetExecutingAssembly().Location.Replace("/", "\\") == current.MainModule.FileName)
                    {
                        return process;
                    }
                }
            }
            return null;

        }
        private static void HandleRunningInstance(Process instance)
        {
            ShowWindowAsync(instance.MainWindowHandle, 1);
            SetForegroundWindow(instance.MainWindowHandle);
        }
        [DllImport("User32.dll")]
        private static extern bool ShowWindowAsync(System.IntPtr hwnd, int cmdShow);
        [DllImport("User.dll")]
        private static extern bool SetForegroundWindow(System.IntPtr hwnd);
        #endregion

        //给条形码变量赋值
        public String SerialNumber
        {
            get 
            {
                return this.serialNum;
            }
            set 
            { 
                this.serialNum =value;
            
            }

        }

        //关闭软件
        private void Button_Click(object sender, RoutedEventArgs e)
        {
            int i=5;
            DateTime dt = System.DateTime.Now;
            string date1 = dt.ToString("yyyyMMdd");

            //判断当天的表是否为空表，是，删除，否则，不删除
            OleDbConnection aConnection = new OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + System.AppDomain.CurrentDomain.BaseDirectory + "Data\\checkDb.mdb");
            string strSql = "Select count(*) from " + date1;//获取表记录数

            try
            {
                aConnection.Open();
                OleDbCommand myCmd = new OleDbCommand(strSql, aConnection);
                i = (int)myCmd.ExecuteScalar();

            }
            catch (Exception ex)
            {
                //MessageBox.Show("ERROR:" + ex.Message);

            }
            finally
            {
                if (aConnection != null)
                    aConnection.Close();

            }

            //当天表为空表,删除表
            if (i <= 0)
            {
                DbOperate del = new DbOperate();

                del.DeleteTable(date1);
            
            }

            var ret = System.Windows.MessageBox.Show("确定退出软件吗？","",MessageBoxButton.YesNo);
            if (ret == MessageBoxResult.Yes)
            {
                //DataProvider.Instance.LoginOut();
                //关闭所有进程
                /*
                try
                {
                    SerialClose();
                }
                catch (Exception)
                {
                    System.Windows.MessageBox.Show("未能正常关闭串口！");
                }
                */

                try
                {
                    Environment.Exit(Environment.ExitCode);
                }
                catch (Exception)
                {
                    //System.Windows.MessageBox.Show("未能正常关闭软件！");
                }
                
            }
        }

        //最小化软件
        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            this.WindowState = System.Windows.WindowState.Minimized;
        }
        //最大化软件
        private void Button_Click_2(object sender, RoutedEventArgs e)
        {
            if (winSizeState == true)//当前软件处于正常大小
            {
                this.WindowState = System.Windows.WindowState.Maximized;
                winSizeState = false;
            }
            else//当前软件处于最大化状态
            {
                this.WindowState = System.Windows.WindowState.Normal;
                winSizeState = true;
            }
            
        }

        //打开配置按钮
        private void config_Click(object sender, RoutedEventArgs e)
        {
            if (cfgOpen == false)
            {
                Thread config = new Thread(new ThreadStart(ConfigShowDialog));

                config.SetApartmentState(ApartmentState.STA);//这个地方必须设置这个STA,否则会报错“调用线程必须为 STA，因为许多 UI 组件都需要。”
                config.IsBackground = true;

                config.Start();
            }
            else
            {
                /*
                if (myCfg.WindowState == WindowState.Minimized)
                    myCfg.WindowState = WindowState.Normal;

                myCfg.Activate();
                 */
            
            }
        }
        public void ConfigShowDialog()
        {
            myCfg = new configForm(this);

            myCfg.Show();
            cfgOpen = true;
            System.Windows.Threading.Dispatcher.Run();//如果去掉这个，会发现启动的窗口显示出来以后会很快就关掉。

            
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            firstCheck.ToolTip = "限20字";
            ////remark.ToolTip = "限20字";

            //数据库不存在，创建数据库
            if (!File.Exists(System.AppDomain.CurrentDomain.BaseDirectory + "Data\\checkDb.mdb"))
            {
                DbOperate testDb = new DbOperate();
                testDb.CreateDb();

            }

            //print.txt不存在，创建默认打印路径
            if (!File.Exists(System.AppDomain.CurrentDomain.BaseDirectory + "Data\\print.txt"))
            {
                string pathStringPrint = System.AppDomain.CurrentDomain.BaseDirectory + "Data\\print.txt";

                try
                {
                    StreamWriter sw = new StreamWriter(pathStringPrint, true, Encoding.GetEncoding("gb2312"));//true:尾部追加

                    sw.WriteLine("2");
                    sw.WriteLine(System.AppDomain.CurrentDomain.BaseDirectory + "Data\\template\\template.xls");

                    sw.Close();
                }
                catch (Exception ex)
                {
                    //什么都不用做
                }

            }

            /*
            //创建存储管理员密码的txt文件
            string pathString = "C:\\temp\\password.txt";
            if (!File.Exists(pathString))//文件不存在，则创建，并写入初始密码：123456
            {
                FileStream fs1 = new FileStream(pathString, FileMode.Create, FileAccess.Write);
                StreamWriter sw = new StreamWriter(fs1);
                MD5_16 myEncryption = new MD5_16();

                sw.WriteLine("123456");
                System.Windows.MessageBox.Show(myEncryption.MD5Encrypt16("123456"));

                sw.Close();
                fs1.Close();

            }
            */

            //创建当天的表
            DateTime dt = System.DateTime.Now;
            string date = dt.ToString("yyyyMMdd");
            OleDbConnection aConnection = new OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + System.AppDomain.CurrentDomain.BaseDirectory + "Data\\checkDb.mdb");
            string strSql = "Select * from " + date;
            //string patientPathString = System.AppDomain.CurrentDomain.BaseDirectory + "Data\\patientInfo.txt";
            //string[] item = new string[6];

            try//判断表是否存在，程序不够严谨（只要判断打开数据库表时出现错误，就归结于表不存在，以后改进）!!
            {
                aConnection.Open();
                OleDbCommand myCmd = new OleDbCommand(strSql, aConnection);
                myCmd.ExecuteNonQuery();

            }
            catch (Exception ex)//表不存在，创建表
            {

                //try
                //{
                //    int i;
                //    FileStream fs1 = new FileStream(patientPathString, FileMode.Open, FileAccess.Read);
                //    StreamReader sr1 = new StreamReader(fs1);

                //    for (i = 1; i < 21; i++)//读取txt文件到21行
                //    {
                //        sr1.ReadLine();
                //    }
                //    for (; i < 32; i = i + 2)
                //    {
                //        item[(i - 21) / 2] = sr1.ReadLine();
                //        sr1.ReadLine();

                //    }

                //    sr1.Close();
                //    fs1.Close();

                //}
                //catch (Exception e1)
                //{
                //    System.Windows.MessageBox.Show("Error6:" + e1.Message);
                //}

                ArrayList headList = new ArrayList();
                DbOperate testDb = new DbOperate();

                headList.Add("医院名称"); headList.Add("科室名称"); headList.Add("仪器型号");
                headList.Add("姓名"); headList.Add("性别"); headList.Add("年龄"); headList.Add("住院号");
                headList.Add("CO"); headList.Add("CO2"); headList.Add("红细胞寿命"); headList.Add("血红蛋白浓度");
                headList.Add("送检医生"); headList.Add("复核医生"); headList.Add("报告医生");
                headList.Add("初步诊断"); 
                headList.Add("时间"); headList.Add("日期"); headList.Add("备注1"); headList.Add("备注2");
                //for (int i = 0; i < 6; i++)
                //{
                //    if (item[i] != "null")
                //        headList.Add(item[i]);
                //}

                testDb.CreateTable(System.AppDomain.CurrentDomain.BaseDirectory + "Data\\checkDb.mdb", date, headList);

            }
            finally
            {
                if (aConnection != null)
                    aConnection.Close();

            }

            //检验医生和复核医生
            string pathString = System.AppDomain.CurrentDomain.BaseDirectory + "Data\\doctor.txt";
            string dcName;
            //string doctorName;

            try
            {
                //FileStream fs1 = new FileStream(pathString, FileMode.Open, FileAccess.ReadWrite);
                StreamReader sr = new StreamReader(pathString, Encoding.GetEncoding("gb2312"));

                while ((dcName = sr.ReadLine()) != null)
                {
                    checkDoctor.Items.Add(dcName);
                    reviewDoctor.Items.Add(dcName);
                }

                sr.Close();
                //fs1.Close();

            }
            catch (Exception ex)
            {
                // System.Windows.MessageBox.Show("ERROR:" + ex.Message);

            }

            //阻止电脑休眠
            SetThreadExecutionState(ES_CONTINUOUS | ES_DISPLAY_REQUIRED | ES_SYSTEM_REQUIRED);

            //聚焦条形码输入框
            tBoxScanBar.Focus();

            //使能或否扫描枪确认按键
            string pathString1 = System.AppDomain.CurrentDomain.BaseDirectory + "Data\\scan.txt";

            try
            {
                StreamReader sr = new StreamReader(pathString1, Encoding.GetEncoding("gb2312"));

                string tmp = sr.ReadLine();

                if (string.Compare(tmp, "0") == 0)
                    scanBarOk.IsEnabled = false;
                else
                    scanBarOk.IsEnabled = true;

                sr.Close();

            }
            catch (Exception ex)
            {
                // System.Windows.MessageBox.Show("ERROR:" + ex.Message);

            }
            //从window load剪接过来的代码
            sex.Items.Add("男");
            sex.Items.Add("女");
            sex.SelectedIndex = 1;
            receiveInfo.Text = "欢迎使用红细胞寿命测定呼气试验仪" + System.Environment.NewLine + "1.确认仪器与软件连接成功；" + System.Environment.NewLine + "2.预热；" + System.Environment.NewLine + "3.将肺泡气袋、本底气袋、倒气袋插入相对应的气嘴处；" + System.Environment.NewLine + "4.按下“输入”键，设置测试者血红蛋白浓度值；" + System.Environment.NewLine + "5.按下“测量”键，开始测量。" + System.Environment.NewLine + "注意:切勿在测量状态下断开USB连接线" + System.Environment.NewLine;
            HosipitalInfoDisplay();
            todayReportDisplay();
            //rbConcentration.Text = "0";
            textboxhb.Text = "0";

            //创建线程，监听连接仪器的串口
            Thread tCom = new Thread(ListenCom);
            tCom.IsBackground = true;
            tCom.Start();           

            Thread tHb = new Thread(listenBGHB);
            tHb.IsBackground = true;
            tHb.Start();


        }

        private void listenBGHB()
        {
            string scanBarCode = null, hosCode = null, url = null;
            string pathStringCom = System.AppDomain.CurrentDomain.BaseDirectory + "Data\\scan.txt";
            //scanBarCode = tBoxScanBar.Text;
            try
            {
                //FileStream fs1 = new FileStream(pathString, FileMode.Open, FileAccess.ReadWrite);
                StreamReader sr = new StreamReader(pathStringCom, Encoding.GetEncoding("gb2312"));

                sr.ReadLine();

                //读入医院代码以及url
                hosCode = sr.ReadLine();
                url = sr.ReadLine();

                sr.Close();
                //fs1.Close();

            }
            catch (Exception ex)
            {
                // System.Windows.MessageBox.Show("ERROR:" + ex.Message);

            }
            string[] args = new string[2];
            args[0] = hosCode;
            //args[1] = scanBarCode;
            try
            {
                while (true)
                {
                    if (websign)
                    {
                        object result=null;
                        bool listisnull = false;
                        for (int i = 0; i < idlist.Length; i++)
                        {
                            if (idlist[i] != null )
                            {

                                listisnull = true;
                                if (timelist[i]!=null)
                                {
                                    args[1] = idlist[i];
                                    result = WebServiceHelper.InvokeWebService(url, "DHCGetHB", args);
                                    if (result.ToString() != "")
                                    {
                                        string pathstringHB = System.AppDomain.CurrentDomain.BaseDirectory + "Data\\HB.txt";
                                        //FileStream fs1 = new FileStream(pathstringHB, FileMode.Create, FileAccess.Write);
                                        //StreamWriter sw1 = new StreamWriter(fs1);
                                        //sw1.Write(result.ToString());
                                        //sw1.Close();
                                        //fs1.Close();
                                        StreamWriter sw = new StreamWriter(pathstringHB, true, Encoding.GetEncoding("gb2312"));//true:尾部追加
                                        sw.WriteLine(result.ToString());
                                        sw.Close();
                                        string RBC = (tmpRBClist[i] * Convert.ToInt32(result) / 100).ToString();
                                        string[] pmts = new string[3];
                                        pmts[0] = idlist[i];
                                        pmts[1] = timelist[i];
                                        pmts[2] = RBC;
                                        object rst = WebServiceHelper.InvokeWebService(url, "DHCUpdateRBC", pmts);
                                        idlist[i] = null;
                                    }
                                }
                                
                            }
                        }
                        if (!listisnull)
                        {
                            websign = false;
                            num = -1;
                        }
                    }
                    Thread.Sleep(2000);
                }
            }
            catch (Exception e31)
            {
                System.Windows.MessageBox.Show("ERROR31:" + e31.Message);
            }
        }

        public void HosipitalInfoDisplay()
        {
            string hosipital;
            string room;
            string device;
            string pathString = System.AppDomain.CurrentDomain.BaseDirectory + "Data\\hosipitalInfo.txt";

            try
            {
                FileStream fs1 = new FileStream(pathString, FileMode.Open, FileAccess.Read);
                StreamReader sr = new StreamReader(fs1);

                hosipital = sr.ReadLine();
                room = sr.ReadLine();
                device = sr.ReadLine();

                hosipitalName.Text = hosipital;
                roomName.Text = room;
                deviceNum.Text = device;

                sr.Close();
                fs1.Close();

            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show("ERROR7:" + ex.Message);

            }

        }

        //private void connDevice_Click(object sender, RoutedEventArgs e)
        //{
        //    SerialOpen();
        //}

        //软件复位
        //private void softwareReset_Click(object sender, RoutedEventArgs e)
        //{
        //    Process p = new Process();
        //    p.StartInfo.FileName = System.AppDomain.CurrentDomain.BaseDirectory + "Seekya.exe";
        //    p.StartInfo.UseShellExecute = false;
        //    p.Start();
        //    System.Windows.Application.Current.Shutdown();  

        //}

        private void todayReportDisplay()
        {
            DateTime dt = System.DateTime.Now;
            string date = dt.ToString("yyyyMMdd");    //获取当天表名
            OleDbConnection aConnection = new OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + System.AppDomain.CurrentDomain.BaseDirectory + "Data\\checkDb.mdb");
            //MessageBox.Show("Select * from " + tableName);
            string querySql = ("Select * from " + date).ToString();

            try
            {
                aConnection.Open();
                OleDbDataAdapter dadapter = new OleDbDataAdapter();
                dadapter.SelectCommand = new OleDbCommand(querySql, aConnection);
                DataSet dSet = new DataSet();

                dadapter.Fill(dSet);

                //获取表的行数
                dbRows = dSet.Tables[0].Rows.Count;

                //为使dataGridView容器，当行数不足以填满容器时，进行补行操作
                if (dbRows < 12)
                {
                    // MessageBox.Show("表中数据的行数为：" + dSet.Tables[0].Rows.Count);
                    int j = dbRows;

                    for (int i = 0; i < (12 - j); i++)                                                                                                                                                                                                                                                                                                                                                                                                                                                             
                    {
                        DataRow dr = dSet.Tables[0].NewRow(); 
                        for (int x = 0; x < 13; x++)
                        {
                            dr[x] = "";//新行的单元格装入空值
                        }
                        dSet.Tables[0].Rows.Add(dr);

                    }

                }
                
                todayReport.DataSource = dSet.Tables[0];

                for (int i = 0; i < 16; i++)    //解除表头（每列头字段）的点中以及排序模式
                    todayReport.Columns[i].SortMode = DataGridViewColumnSortMode.NotSortable;
                //   todayReport.ColumnHeadersDefaultCellStyle.BackColor = System.Drawing.Color.FromArgb(192,0,0);

            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show("Error8:" + ex.Message);

            }
            finally
            {
                if (aConnection != null)
                {
                    aConnection.Close();

                }

                //把当前的cell定位到最新的记录
                if (dbRows!=0)
                    todayReport.CurrentCell = todayReport.Rows[dbRows - 1].Cells[0];

            }

        }

        private void dbManage_Click(object sender, RoutedEventArgs e)
        {
            if (dbOpen == false)
            {
                Thread db = new Thread(new ThreadStart(DbDialogShow));

                db.SetApartmentState(ApartmentState.STA);//这个地方必须设置这个STA,否则会报错“调用线程必须为 STA，因为许多 UI 组件都需要。”
                db.IsBackground = true;

                db.Start();
            }
            else 
            {

                /*
                if (myDb.WindowState == WindowState.Minimized)
                    myDb.WindowState = WindowState.Normal;

                myDb.Activate();
                 * */
            }
            
        }
        private void DbDialogShow()
        {
            myDb = new dbManager(this);

            myDb.Show();
            dbOpen = true;
            System.Windows.Threading.Dispatcher.Run();//如果去掉这个，会发现启动的窗口显示出来以后会很快就关掉。
        
        }
        private void config_ContextMenuClosing(object sender, ContextMenuEventArgs e)
        {

        }
        //拖动串口功能
        private void DragWindow(object sender, MouseButtonEventArgs e)
        {
            DragMove();

        }

        private void printReport_Click(object sender, RoutedEventArgs e)
        {
            PrintReport print = new PrintReport();

            if (todayReport.CurrentCell.RowIndex < dbRows)
            {

                int row = todayReport.CurrentCell.RowIndex;   //定位到最后一条记录

                string[] userDefine = { "", "", "", "", "", "" };
                int i;

                //bool direct = false;//打印方式标志，true:直接打印，false：手动打印
                //string pathString = System.AppDomain.CurrentDomain.BaseDirectory + "Data\\print.txt";//读取打印方式的txt文件

                //try
                //{
                //    FileStream fs1 = new FileStream(pathString, FileMode.Open, FileAccess.Read);
                //    StreamReader sr1 = new StreamReader(fs1);

                //    if (sr1.ReadLine() == "1")
                //        direct = true;
                //    else
                //        direct = false;

                //    sr1.Close();
                //    fs1.Close();

                //}
                //catch (Exception ex)
                //{
                //    //MessageBox.Show("Error:" + ex.Message);
                //}

                //先获取CO浓度和血红蛋白浓度
                string co = todayReport.Rows[row].Cells[7].Value.ToString().Trim();
                string hb = todayReport.Rows[row].Cells[10].Value.ToString().Trim();
                string rbc = todayReport.Rows[row].Cells[9].Value.ToString();
                bool sign = false;

                string hospital = todayReport.Rows[row].Cells[0].Value.ToString();
                string department = todayReport.Rows[row].Cells[1].Value.ToString();
                string instrumentType = todayReport.Rows[row].Cells[2].Value.ToString();
                string name = todayReport.Rows[row].Cells[3].Value.ToString();
                string gender = todayReport.Rows[row].Cells[4].Value.ToString();
                string age = todayReport.Rows[row].Cells[5].Value.ToString();
                string id = todayReport.Rows[row].Cells[6].Value.ToString();
                string co2 = todayReport.Rows[row].Cells[8].Value.ToString();
                string submitDoctor = todayReport.Rows[row].Cells[11].Value.ToString();
                string checkDoctor = todayReport.Rows[row].Cells[12].Value.ToString();
                string reportDoctor = todayReport.Rows[row].Cells[13].Value.ToString();
                string firstVisit = todayReport.Rows[row].Cells[14].Value.ToString();
                string reportTime = todayReport.Rows[row].Cells[15].Value.ToString();
                string testDateLine = todayReport.Rows[row].Cells[16].Value.ToString();
                string remark1 = todayReport.Rows[row].Cells[17].Value.ToString();
                string remark2 = todayReport.Rows[row].Cells[18].Value.ToString();
                //判断血红蛋白浓度是否有效
                if (int.Parse(hb) == 0)
                {
                    hbInput t = new hbInput();

                    t.Owner = this;

                    t.ShowDialog();

                    hb = rbcon;

                    //红细胞寿命换算
                    rbc = ((int)(1.38 * int.Parse(hb) / float.Parse(co))).ToString();
                    sign = true;
                    string date1 = testDateLine.Substring(0, 4) + testDateLine.Substring(5, 2) + testDateLine.Substring(8, 2);
                    string datetime2 = reportTime.Substring(0, 2) + reportTime.Substring(3, 2) + reportTime.Substring(6, 2);

                    string excelName = System.AppDomain.CurrentDomain.BaseDirectory + "Data\\Template\\" + name + "(" + date1 + datetime2 + ")" + ".xls";
                    Open(excelName);
                    Excel.Worksheet ws = (Excel.Worksheet)app.ActiveSheet;

                    DataTable dataTable = new DataTable();
                    dataTable.Columns.Add("hb", typeof(string));
                    dataTable.Columns.Add("rbc", typeof(string));
                    DataRow dr = dataTable.NewRow();
                    dr["hb"] = hb;
                    dr["rbc"] = rbc;
                    dataTable.Rows.Add(dr);
                    app.Goto("rbc");
                    app.ActiveCell.FormulaR1C1 = rbc;
                    app.Goto("hb");
                    app.ActiveCell.FormulaR1C1 = hb;
                    wb.SaveCopyAs(excelName);
                    //wb.Close(Type.Missing, Type.Missing, Type.Missing);      //(W)注意这里有注释掉
                    //wbs.Close();                                             //(W)注意这里有注释掉
                    //app.Quit();
                    wb = null;
                    wbs = null;
                    //app = null;
                    GC.Collect();
                    PublicMethod.Kill(app);
                }



                try
                {
                    for (i = 18; i < 24; i++)
                    {
                        userDefine[i - 18] = todayReport.Rows[row].Cells[i].Value.ToString();

                    }

                }
                catch { }

                try
                {
                    if (prtmd == true)   //直接打印
                        print.ReportPrintDirect(name, gender, age, id, instrumentType, submitDoctor, firstVisit, hb, hospital, rbc, co, co2, testDateLine, department, userDefine[0], userDefine[1], userDefine[2], userDefine[3], userDefine[4], userDefine[5], checkDoctor, reportDoctor, reportTime, remark1, remark2);
                    else    //手动打印
                        print.ReportPrintHand(name, gender, age, id, instrumentType, submitDoctor, firstVisit, hb, hospital, rbc, co, co2, testDateLine, department, userDefine[0], userDefine[1], userDefine[2], userDefine[3], userDefine[4], userDefine[5], checkDoctor, reportDoctor, reportTime, remark1, remark2);
                }
                catch (Exception eee)
                {
                    System.Windows.MessageBox.Show("ERROR25:" + eee.Message);
                }

                if (sign == true)
                {
                    //修改数据库未存有血红蛋白浓度的检验数据
                    DbOperate test = new DbOperate();
                    //更改记录
                    test.ModifyRecord(testDateLine.Substring(0, 4) + testDateLine.Substring(5, 2) + testDateLine.Substring(8, 2), reportTime, hospital, department, instrumentType, name, gender, age, id, co, co2, rbc, hb, submitDoctor, checkDoctor, reportDoctor, firstVisit, reportTime, testDateLine, remark1, remark2);

                }

                //刷新当天检验报告
                todayReportDisplay();

            }
        }
     
        //打开打印模板，让用户修改模板
        private void printModel_Click(object sender, RoutedEventArgs e)
        {
            PrintReport myPrint = new PrintReport();

            Process.Start(System.AppDomain.CurrentDomain.BaseDirectory + "Data\\out.xls");

        }


        private void help_Click(object sender, RoutedEventArgs e)
        {
            //打开帮助文档
            System.Diagnostics.Process.Start(System.AppDomain.CurrentDomain.BaseDirectory + "Data\\help.CHM");

        }

        private void button1_Click(object sender, RoutedEventArgs e)
        {
            /*
            Thread hl7 = new Thread(new ThreadStart(showScan));

            hl7.SetApartmentState(ApartmentState.STA);
            hl7.IsBackground = true;

            hl7.Start();
            */
            HL7 hl7 = new HL7(this);

            hl7.ShowDialog();
        }
        //显示条形码扫描界面
        private void showScan()
        {
            HL7 myHL7 = new HL7(this);

            myHL7.ShowDialog();

            System.Windows.Threading.Dispatcher.Run();//如果去掉这个，会发现启动的窗口显示出来以后会很快就关掉。
        
        }


        /**************************************************
         **描述：HL7部分的代码
         **时间：2017-7-27
         ***************************************************/
        #region 变量
        // 申明变量
        private TcpClient tcpClient = null;
        private NetworkStream networkStream = null;
        private BinaryReader reader;
        private BinaryWriter writer;

        //
        Socket s = new Socket(AddressFamily.InterNetwork, SocketType.Stream, ProtocolType.Tcp);
        
        // 申明委托
        // 显示消息
        //private delegate void ShowMessage(string str);
        //private ShowMessage showMessageCallback;

        // 显示状态
        //private delegate void ShowStatus(string str);
        //private ShowStatus showStatusCallBack;
        
        // 清空消息
        //private delegate void ResetMessage();
        //private ResetMessage resetMessageCallBack;

        #endregion 

        //HL7解析变量
        private static XmlDocument _xmlDoc;

        String serverIP = "";
        String serverPort="";

        //连接服务器标志
        bool connSign = false;//false:表示未连接，true：已连接过，包括连接后又断开

        //存储接收到的数据
        //byte[] res = new byte[1024];

        int qryOroru = 0;//1:代表发送QRY 2：表示发送ORU

        //QRY
        String qry = "MSH|^~\\&|SeekyaRBCS1.1.0||LIS||||QRY^R02||P|2.3.1"+"\r"+"QRD||R|I|||10|RD|120026|RES"+"\r"+"\x1C\r";

        //ORU
        String oru = "MSH|^~\\&|SeekyaRBCS1.1.0||LIS||||ORU^R01||P|2.3.1"+"\r"+ "PID||120026|||LiXiao||18|M" + "\r" +"PV1|医院名称|科室名称|仪器型号|送检医生|初步诊断|备注|血红蛋白浓度|红细胞寿命|CO|CO2|时间|日期"+"\r"+"\x1C\r";

        /*
        public frmSyncTCPClient()
        {
            InitializeComponent();

            #region 实例化委托
            // 显示消息
            showMessageCallback = new ShowMessage(showMessage);

            // 显示状态
            showStatusCallBack = new ShowStatus(showStatus);       

            // 重置消息
            resetMessageCallBack = new ResetMessage(resetMessage);
            #endregion               
        }
        
        #region 定义回调函数

        // 显示消息
        private void showMessage(string str)
        {
            lstbxMessageView.Items.Add(tcpClient.Client.RemoteEndPoint);
            lstbxMessageView.Items.Add(str);
            lstbxMessageView.TopIndex = lstbxMessageView.Items.Count - 1;
        }

        // 显示状态
        private void showStatus(string str)
        {
            toolStripStatusInfo.Text = str;
        }
         
        // 清空消息
        private void resetMessage()
        {
            tbxMessage.Text = "";
            tbxMessage.Focus();
        }

        #endregion 
        */
        #region 点击事件方法
       /*
        * private void btnConnect_Click(object sender, EventArgs e)
        {
            // 通过一个线程发起请求,多线程
            Thread connectThread = new Thread(ConnectToServer);
            connectThread.Start();
        }
        */
        // 连接服务器方法,建立连接的过程
        public void ConnectToServer()
        {
            
            //try
            //{
                if (serverIP == string.Empty || serverPort == string.Empty)
                {
                    System.Windows.MessageBox.Show("请先输入服务器的IP地址和端口号");

                    return;
                }

                IPAddress ipaddress = IPAddress.Parse(serverIP);
                //tcpClient = new TcpClient();             
                //tcpClient.Connect(ipaddress, int.Parse(serverPort));   

                if (connSign == false)
                {
                    s.Connect(ipaddress, int.Parse(serverPort));

                    //客户端定时接收
                    //Control.CheckForIllegalCrossThreadCalls = false;//WinForm

                    aTimer = new System.Timers.Timer(); //实例化定时器
                    aTimer.Enabled = true;
                    Thread thread1 = new Thread(TimerMange);
                    thread1.IsBackground = true;
                    thread1.Start();

                }
                else
                {
                    s = new Socket(AddressFamily.InterNetwork, SocketType.Stream, ProtocolType.Tcp);
                    s.Connect(ipaddress, int.Parse(serverPort));

                    aTimer.Enabled = true;


                }
                
                // 延时操作
                //Thread.Sleep(1000);
                //if (tcpClient != null)
                //{
                    //networkStream = tcpClient.GetStream();
                    //reader = new BinaryReader(networkStream);
                    //writer =new BinaryWriter(networkStream);
                //}

                
                connSign = true;
                System.Windows.MessageBox.Show("远程连接成功");
                
            //}
            //catch
            //{
             //   System.Windows.MessageBox.Show("远程连接失败,请确保输入的服务器IP与端口号无误");
            //    Thread.Sleep(1000);
            //}
        }
        //
        private void TimerMange()
        {
            aTimer.Elapsed += new ElapsedEventHandler(receiveMessage);//添加定时事件触发
            aTimer.Interval = 100;

        }

        //HL7转换为XML
        public static XmlDocument ConvertToXmlObject(string sHL7)
        {
            _xmlDoc = CreateXmlDoc();

            //把HL7分成段
            string[] sHL7Lines = sHL7.Split('\r');

            //去掉XML的关键字
            for (int i = 0; i < sHL7Lines.Length; i++)
            {

                sHL7Lines[i] = sHL7Lines[i].Replace(@"[^ -~]", "");
                sHL7Lines[i] = sHL7Lines[i].Replace("\v", "");
                sHL7Lines[i] = sHL7Lines[i].Replace("\x1C", "");

            }

            for (int i = 0; i < sHL7Lines.Length; i++)
            {
                // 判断是否空行
                if (sHL7Lines[i] != string.Empty)
                {
                    string sHL7Line = sHL7Lines[i];

                    //通过/r 或/n 回车符分隔
                    string[] sFields = GetMessgeFields(sHL7Line);

                    // 为段（一行）创建第一级节点
                    XmlElement el = _xmlDoc.CreateElement(sFields[0]);
                    _xmlDoc.DocumentElement.AppendChild(el);

                    // 循环每一行
                    for (int a = 0; a < sFields.Length; a++)
                    {
                        // 为字段创建第二级节点
                        XmlElement fieldEl = _xmlDoc.CreateElement(sFields[0] + "." + a.ToString());

                        //是否包括HL7的连接符
                        if (sFields[a] != @"^~\&")
                        {//0:如果这一行有任何分隔符

                            //通过~分隔
                            string[] sComponents = GetRepetitions(sFields[a]);
                            if (sComponents.Length > 1)
                            {//1:如果可以分隔
                                for (int b = 0; b < sComponents.Length; b++)
                                {

                                    XmlElement componentEl = _xmlDoc.CreateElement(sFields[0] + "." + a.ToString() + "." + b.ToString());

                                    //通过&分隔 
                                    string[] subComponents = GetSubComponents(sComponents[b]);
                                    if (subComponents.Length > 1)
                                    {//2.如果有字组，一般是没有的。。。
                                        for (int c = 0; c < subComponents.Length; c++)
                                        {
                                            //修改了一个错误
                                            string[] subComponentRepetitions = GetComponents(subComponents[c]);
                                            if (subComponentRepetitions.Length > 1)
                                            {
                                                for (int d = 0; d < subComponentRepetitions.Length; d++)
                                                {
                                                    XmlElement subComponentRepEl = _xmlDoc.CreateElement(sFields[0] + "." + a.ToString() + "." + b.ToString() + "." + c.ToString() + "." + d.ToString());
                                                    subComponentRepEl.InnerText = subComponentRepetitions[d];
                                                    componentEl.AppendChild(subComponentRepEl);
                                                }
                                            }
                                            else
                                            {
                                                XmlElement subComponentEl = _xmlDoc.CreateElement(sFields[0] + "." + a.ToString() + "." + b.ToString() + "." + c.ToString());
                                                subComponentEl.InnerText = subComponents[c];
                                                componentEl.AppendChild(subComponentEl);

                                            }
                                        }
                                        fieldEl.AppendChild(componentEl);
                                    }
                                    else
                                    {//2.如果没有字组了，一般是没有的。。。
                                        string[] sRepetitions = GetComponents(sComponents[b]);
                                        if (sRepetitions.Length > 1)
                                        {
                                            XmlElement repetitionEl = null;
                                            for (int c = 0; c < sRepetitions.Length; c++)
                                            {
                                                repetitionEl = _xmlDoc.CreateElement(sFields[0] + "." + a.ToString() + "." + b.ToString() + "." + c.ToString());
                                                repetitionEl.InnerText = sRepetitions[c];
                                                componentEl.AppendChild(repetitionEl);
                                            }
                                            fieldEl.AppendChild(componentEl);
                                            el.AppendChild(fieldEl);
                                        }
                                        else
                                        {
                                            componentEl.InnerText = sComponents[b];
                                            fieldEl.AppendChild(componentEl);
                                            el.AppendChild(fieldEl);
                                        }
                                    }
                                }
                                el.AppendChild(fieldEl);
                            }
                            else
                            {//1:如果不可以分隔，可以直接写节点值了。
                                fieldEl.InnerText = sFields[a];
                                el.AppendChild(fieldEl);
                            }

                        }
                        else
                        {//0:如果不可以分隔，可以直接写节点值了。
                            fieldEl.InnerText = sFields[a];
                            el.AppendChild(fieldEl);
                        }
                    }
                }
            }

            return _xmlDoc;
        }
        /// <summary>
        /// 通过|分隔 字段
        /// </summary>
        /// <param name="s"></param>
        /// <returns></returns>
        private static string[] GetMessgeFields(string s)
        {
            return s.Split('|');
        }

        /// <summary>
        /// 通过^分隔 组字段
        /// </summary>
        /// <param name="s"></param>
        /// <returns></returns>
        private static string[] GetComponents(string s)
        {
            return s.Split('^');
        }

        /// <summary>
        /// 通过&分隔 子分组组字段
        /// </summary>
        /// <param name="s"></param>
        /// <returns></returns>
        private static string[] GetSubComponents(string s)
        {
            return s.Split('&');
        }

        /// <summary>
        /// 通过~分隔 重复
        /// </summary>
        /// <param name="s"></param>
        /// <returns></returns>
        private static string[] GetRepetitions(string s)
        {
            return s.Split('~');
        }

        /// <summary>
        /// 创建XML对象
        /// </summary>
        /// <returns></returns>
        private static XmlDocument CreateXmlDoc()
        {
            XmlDocument output = new XmlDocument();
            XmlElement rootNode = output.CreateElement("HL7Message");
            output.AppendChild(rootNode);
            return output;
        }

        public static string GetText(XmlDocument xmlObject, string path)
        {
            XmlNode node = xmlObject.DocumentElement.SelectSingleNode(path);
            if (node != null)
            {
                return node.InnerText;
            }
            else
            {
                return null;
            }
        }

        public static string GetText(XmlDocument xmlObject, string path, int index)
        {
            XmlNodeList nodes = xmlObject.DocumentElement.SelectNodes(path);
            if (index <= nodes.Count)
            {
                return nodes[index].InnerText;
            }
            else
            {
                return null;
            }

        }

        public static String[] GetTexts(XmlDocument xmlObject, string path)
        {
            XmlNodeList nodes = xmlObject.DocumentElement.SelectNodes(path);
            String[] arr = new String[nodes.Count];
            int index = 0;
            foreach (XmlNode node in nodes)
            {
                arr[index++] = node.InnerText;
            }
            return arr;

        }
        // 接受消息
        private void receiveMessage(object sender, EventArgs e)
        {
           // System.Windows.MessageBox.Show("1");

           Byte[] res = new Byte[1024];
           try
           {
                //测试
                //Thread.Sleep(1000);

                int receiveLength = s.Receive(res, res.Length, SocketFlags.None);

                if (receiveLength > 0)
                {

                     Encoding gb = System.Text.Encoding.GetEncoding("gb2312");

                    String receivemessage = gb.GetString(res);
                    //s.Send(Encoding.ASCII.GetBytes(abc));

                    //string receivemessage = reader.ReadString().Trim();

                    //System.Windows.MessageBox.Show("接收：" + receivemessage);


                        if (String.Compare(receivemessage, "AA") == 0)//接收到AA
                        {
                            Thread sendThread = new Thread(SendMessage);

                            sendThread.Start("BB");                   

                        }
                        else if (String.Compare(receivemessage, "BB") == 0)//接收到BB
                        {

                            if (qryOroru == 1)//发送QRY
                            {

                                String qryTemp = "";
                                //加入校验位
                                Byte[] arr2 = gb.GetBytes(qry);

                                Byte[] arr1 = new Byte[2];//{ 0X0B, (Byte)(arr2.Length % 256) };
                                arr1[0] = 0X0B;
                                arr1[1] = (Byte)((arr2.Length) % 256);

                                Byte[] arr = new Byte[arr1.Length + arr2.Length];

                                //获得需要发送的消息，加上了校验位
                                arr1.CopyTo(arr, 0);
                                arr2.CopyTo(arr, arr1.Length);

                                qryTemp = gb.GetString(arr);

                                //发送qry
                                Thread sendQry = new Thread(SendMessage);
                                sendQry.IsBackground = true;

                                sendQry.Start(qryTemp);

                            }
                            else if (qryOroru == 2)//发送ORU
                            {
                                String oruTemp = "";

                                //加入校验位       
                                Byte[] arr2 = gb.GetBytes(oru);

                                Byte[] arr1 = { 0X0B, (Byte)((arr2.Length) % 256) };
                                Byte[] arr = new Byte[arr1.Length + arr2.Length];

                                //获得需要发送的消息，加上了校验位
                                arr1.CopyTo(arr, 0);
                                arr2.CopyTo(arr, arr1.Length);

                                oruTemp = gb.GetString(arr);

                                Thread sendOru = new Thread(SendMessage);
                                sendOru.IsBackground = true;

                                sendOru.Start(oruTemp);

                            }

                        }
                        else if (String.Compare(receivemessage, "00") == 0)
                        {

                            //对方接受成功，不需要应答
                            qryOroru = 0;

                        }
                        else if (String.Compare(receivemessage, "FF") == 0)
                        {

                            //对方接受失败，重复发送消息
                            if (qryOroru == 1)//发送QRY
                            {
                                Thread sendQry = new Thread(SendMessage);
                                sendQry.IsBackground = true;

                                sendQry.Start(qry);

                            }
                            else if (qryOroru == 2)//发送ORU
                            {
                                Thread sendOru = new Thread(SendMessage);
                                sendOru.IsBackground = true;

                                sendOru.Start(oru);

                            }

                        }
                        else//接收到DSR
                        {
                            
                            Byte[] arr = gb.GetBytes(receivemessage);
                            //存储arr中消息的前两位后的数据，即除去消息的“\v”和检验位
                            Byte[] arrTemp=new Byte[receiveLength-2];

                            //System.Windows.MessageBox.Show((receiveLength).ToString());

                            //lstbxMessageView.Invoke(showMessageCallback, receivemessage);
                        
                            if (res[1] == ((receiveLength - 2) % 256))//检验成功
                            {
                                //把接收到的字符串剪切掉校验位
                                Array.Copy(arr, 2, arrTemp, 0, (receiveLength - 2));
                                receivemessage = gb.GetString(arrTemp);

                                //解析DSR
                                XmlDocument xmlObject = ConvertToXmlObject(receivemessage);

                                outputID(GetText(xmlObject, "PID/PID.2", 0));    //住院号
                                outputName(GetText(xmlObject, "PID/PID.5", 0));    //姓名
                                outputAge(GetText(xmlObject, "PID/PID.7", 0));    //年龄
                                outputSex(GetText(xmlObject, "PID/PID.8", 0));    //性别

                                outputSendDoctor(GetText(xmlObject, "PV1/PV1.1", 0));    //送检医生
                                outputFirstCheck(GetText(xmlObject, "PV1/PV1.2", 0));    //初步诊断
                                //outputRemark(GetText(xmlObject, "PV1/PV1.3", 0));    //备注
                                outputRBConcentration(GetText(xmlObject, "PV1/PV1.4", 0));    //血红蛋白浓度

                                //回复接收成功
                                Thread send00 = new Thread(SendMessage);

                                send00.IsBackground = true;

                                send00.Start("00");
                            }
                            else
                            {
                                //回复接收失败
                                Thread sendFF = new Thread(SendMessage);

                                sendFF.IsBackground = true;

                                sendFF.Start("FF");
                            }
                        }
                }
            }
            catch
            {
                //s.Shutdown(SocketShutdown.Both);
                //s.Close();

                //aTimer.Stop();
                //aTimer.Enabled = false;

                //System.Windows.MessageBox.Show("出错");
            }
       

        }


        // 断开连接
        public void DisconnectToServer()
        {
            try
            {
                /*
                 * if (reader != null)
                {
                    reader.Close();
                }
                if (writer != null)
                {
                    writer.Close();
                }
                if (tcpClient != null)
                {
                    // 断开连接
                    tcpClient.Close();
                }
                */
                s.Close();

                aTimer.Stop();
                aTimer.Enabled = false;

                System.Windows.MessageBox.Show("连接已断开");
            }
            catch 
            {
                System.Windows.MessageBox.Show("断开连接失败");
            
            }
        }

        // 关闭窗口
        private void btnClose_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        // 发送消息
        private void btnSend_Click(object sender, EventArgs e)
        {
            Thread sendThread = new Thread(SendMessage);
            sendThread.Start(qry);
        }

        private void SendMessage(object state)
        {
            Encoding gb = System.Text.Encoding.GetEncoding("gb2312"); 

            try
            {
                s.Send(gb.GetBytes(state.ToString()));
                Thread.Sleep(5000);
                //writer.Flush();

            }
            catch
            {
                /*
                if (reader != null)
                {
                    reader.Close();
                }
                if (writer != null)
                {
                    writer.Close();
                }
                if (tcpClient != null)
                {
                    tcpClient.Close();
                }
                */

                //s.Shutdown(SocketShutdown.Both);
                //s.Close();

                //关闭计数器
                //aTimer.Stop();
                //aTimer.Enabled = false;

                System.Windows.MessageBox.Show("TCP发送出错");
            }
        }

        #endregion

        //扫描条形码后，先发送AA，对方空闲，发送QRY
        public void QRYAA()
        {
            qry = "MSH|^~\\&|SeekyaRBCS1.1.0||LIS||||QRY^R02||P|2.3.1" + "\r" + "QRD||R|I|||10|RD|"+SerialNumber+"|RES" + "\r" + "\x1C\r";
            
            qryOroru = 1;//需要发送QRY

            //发送AA
            Thread sendAA = new Thread(SendMessage);
            sendAA.IsBackground = true;

            sendAA.Start("AA");
        }

        public void ORUAA(object obj)
        {

            String[] temp = obj.ToString().Split('@');

            oru = "MSH|^~\\&|SeekyaRBCS1.1.0||LIS||||ORU^R01||P|2.3.1" + "\r" + "PID||" + temp[0] + "|||" + temp[1] + "||" + temp[2] + "|" + temp[3] + "\r" + "PV1|" + temp[4] + "|" + temp[5] + "|" + temp[6] + "|" + temp[7] + "|" + temp[8] + "|" + temp[9] + "|" + temp[10] + "|" + temp[11] +"|"+temp[12]+"|"+temp[13]+"|"+temp[14]+ "|"+temp[15]+"\r" + "\x1C\r";

            qryOroru = 2;//需要发送ORU

            //发送AA
            Thread sendAA = new Thread(SendMessage);
            sendAA.IsBackground = true;

            sendAA.Start("AA");
        }
        //serverIP变量
        public String ServerIP
        {
            get
            {
                return serverIP;
            }
            set
            {
                this.serverIP = value ;
            
            }
        }
        //serverPort变量
        public String ServerPort
        {
            get
            {
                return serverPort;

            }
            set 
            {
                this.serverPort = value;

            }
        
        }
        //控件跨线程访问
        private delegate void outputDelegate(string msg);

        //姓名textBox
        private void outputName(string msg)
        {
            this.name.Dispatcher.Invoke(new outputDelegate(outputAction1), msg);
        }

        private void outputAction1(string msg)
        {
            this.name.Text=msg;
            //this.name.AppendText("\n");
        }

        //住院号textBox
        private void outputID(string msg)
        {
            this.id.Dispatcher.Invoke(new outputDelegate(outputAction2), msg);
        }

        private void outputAction2(string msg)
        {
            this.id.Text=msg;
            //this.id.AppendText("\n");
        }

        //年龄textBox
        private void outputAge(string msg)
        {
            this.age.Dispatcher.Invoke(new outputDelegate(outputAction3), msg);
        }

        private void outputAction3(string msg)
        {
            this.age.Text=msg;
            //this.age.AppendText("\n");
        }

        //性别textBox
        private void outputSex(string msg)
        {
            this.sex.Dispatcher.Invoke(new outputDelegate(outputAction4), msg);
        }

        private void outputAction4(string msg)
        {
            this.sex.Text = msg ;
            //this.sex.AppendText("\n");
        }

        //送检医生textBox
        private void outputSendDoctor(string msg)
        {
            this.sendDoctor.Dispatcher.Invoke(new outputDelegate(outputAction5), msg);
        }

        private void outputAction5(string msg)
        {
            this.sendDoctor.Text=msg;
            //this.sendDoctor.AppendText("\n");
        }

        //初步诊断textBox
        private void outputFirstCheck(string msg)
        {
            this.firstCheck.Dispatcher.Invoke(new outputDelegate(outputAction6), msg);
        }

        private void outputAction6(string msg)
        {
            this.firstCheck.Text=msg;
            //this.firstCheck.AppendText("\n");
        }



        //血红蛋白浓度textBox
        private void outputRBConcentration(string msg)
        {
            //this.rbConcentration.Dispatcher.Invoke(new outputDelegate(outputAction8), msg);
            this.textboxhb.Dispatcher.Invoke(new outputDelegate(outputAction8), msg);
        }

        private void outputAction8(string msg)
        {
            //this.rbConcentration.Text=msg;
            //this.rbConcentration.AppendText("\n");
            this.textboxhb.Text = msg;
        }

        private void button2_Click(object sender, RoutedEventArgs e)
        {
            todayReportDisplay();
        }

        //todayReport跨线程操作
        public void UpdateTodayReport()
        {
            if (todayReport.InvokeRequired)
            {
                // 当一个控件的InvokeRequired属性值为真时，说明有一个创建它以外的线程想访问它
                Action<string> actionDelegate = (x) => { todayReportDisplay(); };
                // 或者
                // Action<string> actionDelegate = delegate(string txt) { this.label2.Text = txt; };
                this.todayReport.Invoke(actionDelegate);
            }
            else
            {
               //非跨线程访问todayReport
            }
        }

        private void button1_Click_1(object sender, RoutedEventArgs e)
        {
            Process.Start(System.AppDomain.CurrentDomain.BaseDirectory + "Data\\zero.txt");
        }
        //刷新主界面的检验记录表
        private delegate void setRichTexBox();

        public void setText()
        {
            if (this.todayReport.InvokeRequired)//等待异步
            {
                setRichTexBox fc = new setRichTexBox(Set);
                this.todayReport.Invoke(fc, new object[] { });
            }
            else
            {
                //什么都不用做
            }
        }
        private void Set()
        {
            todayReportDisplay();
        }

        //刷新主界面医院信息
        private delegate void SetHosipitalInfo(string s);

        public void RefresHosipitalInfo()
        {
            //读取医院信息
            string pathString = System.AppDomain.CurrentDomain.BaseDirectory + "Data\\hosipitalInfo.txt";
            string hn = null, rn = null, dn = null;

            try
            {
                FileStream fs1 = new FileStream(pathString, FileMode.Open, FileAccess.Read);
                StreamReader sr = new StreamReader(fs1);

                hn = sr.ReadLine();
                rn = sr.ReadLine();
                dn = sr.ReadLine();

                sr.Close();
                fs1.Close();

            }
            catch (Exception ex)
            {
                // System.Windows.MessageBox.Show("ERROR:" + ex.Message);

            }
            this.hosipitalName.Dispatcher.Invoke(new SetHosipitalInfo(SetHosipitalName), hn);
            this.roomName.Dispatcher.Invoke(new SetHosipitalInfo(SetRoomName), rn);
            this.deviceNum.Dispatcher.Invoke(new SetHosipitalInfo(SetDeviceNum), dn);
        }
        private void SetHosipitalName(string s)
        {
            hosipitalName.Text = s;
        }
        private void SetRoomName(string s)
        {
            roomName.Text = s;
        }
        private void SetDeviceNum(string s)
        {
            deviceNum.Text = s;
        }

        //刷新主界面的医生名字信息
        private delegate void SetDoctor();

        public void RefreshDoctor()
        {
            this.checkDoctor.Dispatcher.Invoke(new SetDoctor(SetCheckDoctor));
            this.checkDoctor.Dispatcher.Invoke(new SetDoctor(SetReviewDoctor));

        }

        private void SetCheckDoctor()
        {
            //检验医生
            string pathString = System.AppDomain.CurrentDomain.BaseDirectory + "Data\\doctor.txt";
            string dcName;
            //string doctorName;

            try
            {
                //FileStream fs1 = new FileStream(pathString, FileMode.Open, FileAccess.ReadWrite);
                StreamReader sr = new StreamReader(pathString, Encoding.GetEncoding("gb2312"));

                checkDoctor.Items.Clear();

                while ((dcName = sr.ReadLine()) != null)
                {
                    checkDoctor.Items.Add(dcName);

                }

                sr.Close();
                //fs1.Close();

            }
            catch (Exception ex)
            {
                // System.Windows.MessageBox.Show("ERROR:" + ex.Message);

            }

        }

        private void SetReviewDoctor()
        {
            //复核医生
            string pathString = System.AppDomain.CurrentDomain.BaseDirectory + "Data\\doctor.txt";
            string dcName;
            //string doctorName;

            try
            {
                //FileStream fs1 = new FileStream(pathString, FileMode.Open, FileAccess.ReadWrite);
                StreamReader sr = new StreamReader(pathString, Encoding.GetEncoding("gb2312"));

                reviewDoctor.Items.Clear();

                while ((dcName = sr.ReadLine()) != null)
                {
                    reviewDoctor.Items.Add(dcName);

                }

                sr.Close();
                //fs1.Close();

            }
            catch (Exception ex)
            {
                // System.Windows.MessageBox.Show("ERROR:" + ex.Message);

            }
        }

        private void btnScan_Click(object sender, RoutedEventArgs e)
        {
             barCode scanBarCode = new barCode(this);

            scanBarCode.ShowDialog();

        }


        private void measure_Click(object sender, RoutedEventArgs e)
        {
            Byte[] temp = new Byte[6];
            //获取接收数据时的系统时间
            DateTime dt1 = System.DateTime.Now;
            string date3 = dt1.ToLocalTime().ToString();

            temp[5] = 0X20;
            temp[4] = 0x00;
            temp[3] = 0x00;
            temp[2] = 0x00;
            temp[1] = 0x00;
            temp[0] = 0x20;

            //写日志
            WriteLog("[" + date3 + "]" + ":" + "200000000020");

            sp.Write(temp, 0, 6);
        }

        private void QC_Click(object sender, RoutedEventArgs e)
        {
            /*
            Thread qc = new Thread(new ThreadStart(qcDialogShow));

            qc.SetApartmentState(ApartmentState.STA);//这个地方必须设置这个STA,否则会报错“调用线程必须为 STA，因为许多 UI 组件都需要。”
            qc.IsBackground = true;

            qc.Start();
             * */
            softwareOperate = true;
            if (qcOpen == false)
                qcDialogShow();
            else
            {
                if (myQC.WindowState == WindowState.Minimized)
                    myQC.WindowState = WindowState.Normal;

                myQC.Activate();
            
            }

        }
        private void qcDialogShow()
        {
            myQC = new QC(this);

            myQC.Show();

            qcOpen = true;
            //System.Windows.Threading.Dispatcher.Run();//如果去掉这个，会发现启动的窗口显示出来以后会很快就关掉。


        }

        //使能扫描按键
        private delegate void SetScanBar();

        public void enableBar()
        {
            this.scanBarOk.Dispatcher.Invoke(new SetScanBar(SetEnableBar));
        }

        public void unenableBar()
        {
            this.scanBarOk.Dispatcher.Invoke(new SetScanBar(SetUnenableBar));
        }

        //使能条形码按键函数
        public void SetEnableBar()
        {
            scanBarOk.IsEnabled = true;

        }

        //不使能条形码按键函数
        public void SetUnenableBar()
        {
            scanBarOk.IsEnabled = false;

        }

        private void scanBarOk_Click(object sender, RoutedEventArgs e)
        {
            string scanBarCode = null, hosCode = null, url = null;
            string pathStringCom = System.AppDomain.CurrentDomain.BaseDirectory + "Data\\scan.txt";

            scanBarCode = tBoxScanBar.Text;

            try
            {
                //FileStream fs1 = new FileStream(pathString, FileMode.Open, FileAccess.ReadWrite);
                StreamReader sr = new StreamReader(pathStringCom, Encoding.GetEncoding("gb2312"));

                sr.ReadLine();

                //读入医院代码以及url
                hosCode = sr.ReadLine();
                url = sr.ReadLine();

                sr.Close();
                //fs1.Close();

            }
            catch (Exception ex)
            {
                // System.Windows.MessageBox.Show("ERROR:" + ex.Message);

            }

            //向后台发送条形码
            string[] args = new string[2];
            args[0] = hosCode;
            args[1] = scanBarCode;
            object result = WebServiceHelper.InvokeWebService(url, "DHCGetXXByLabno", args);

            //解析获取到的患者信息
            XmlDocument doc = new XmlDocument();

            doc.LoadXml(result.ToString());
            
            XmlElement root = null;
            root = doc.DocumentElement;
            XmlNodeList listNodes = null;

            //住院号
            listNodes = root.SelectNodes("/DHCLISTOHXBSMLIST/DHCLISTOHXBSM/zyh");

            foreach (XmlNode node in listNodes)
            {
                id.Text = node.InnerText;
            }

            //姓名
            listNodes = root.SelectNodes("/DHCLISTOHXBSMLIST/DHCLISTOHXBSM/patname");

            foreach (XmlNode node in listNodes)
            {
                name.Text = node.InnerText;
            }

            //性别
            listNodes = root.SelectNodes("/DHCLISTOHXBSMLIST/DHCLISTOHXBSM/Sex");

            foreach (XmlNode node in listNodes)
            {
                sex.Text = node.InnerText;
            }

            if (String.Compare(sex.Text, "M") == 0)
                sex.Text = "男";
            else
                sex.Text = "女";

            //年龄
            listNodes = root.SelectNodes("/DHCLISTOHXBSMLIST/DHCLISTOHXBSM/age");

            foreach (XmlNode node in listNodes)
            {
                age.Text = node.InnerText;
            }

            //报告医生
            listNodes = root.SelectNodes("/DHCLISTOHXBSMLIST/DHCLISTOHXBSM/bgdoctor");

            foreach (XmlNode node in listNodes)
            {
                checkDoctor.Text = node.InnerText;
            }

            //复核医生
            listNodes = root.SelectNodes("/DHCLISTOHXBSMLIST/DHCLISTOHXBSM/fhdoctor");

            foreach (XmlNode node in listNodes)
            {
                reviewDoctor.Text = node.InnerText;
            }

            //送检医生
            listNodes = root.SelectNodes("/DHCLISTOHXBSMLIST/DHCLISTOHXBSM/sjdoctor");

            foreach (XmlNode node in listNodes)
            {
                sendDoctor.Text = node.InnerText;
            }

            //血红蛋白浓度
            listNodes = root.SelectNodes("/DHCLISTOHXBSMLIST/res/DHCLISTOHXBSMRES/result");

            foreach (XmlNode node in listNodes)
            {
                //rbConcentration.Text = node.InnerText;
                textboxhb.Text = node.InnerText;
            }

            //初诊
            listNodes = root.SelectNodes("/DHCLISTOHXBSMLIST/DHCLISTOHXBSM/cz");

            foreach (XmlNode node in listNodes)
            {
                firstCheck.Text = node.InnerText;
            }
            string pathString = System.AppDomain.CurrentDomain.BaseDirectory + "Data\\patientInfo.txt";
            FileStream fs1 = new FileStream(pathString, FileMode.Open, FileAccess.Read);
            StreamReader sr1 = new StreamReader(fs1);
            for(int i=0;i<12;i++)
            {
                string nd = sr1.ReadLine();
                propts[i] = nd;
                if (string.Compare(nd,"NULL")!=0)
                {
                    listNodes = root.SelectNodes("/DHCLISTOHXBSMLIST/DHCLISTOHXBSM/"+nd);
                    foreach (XmlNode node in listNodes)
                    {
                        values[i] = node.InnerText;
                    }
                }
            }
            websign = true;
            wsn = true;
            num++;
            idlist[num] = scanBarCode;


        }

        private void Window_Closed(object sender, EventArgs e)
        {
            int i = 5;
            DateTime dt = System.DateTime.Now;
            string date1 = dt.ToString("yyyyMMdd");

            //判断当天的表是否为空表，是，删除，否则，不删除
            OleDbConnection aConnection = new OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + System.AppDomain.CurrentDomain.BaseDirectory + "Data\\checkDb.mdb");
            string strSql = "Select count(*) from " + date1;//获取表记录数

            try
            {
                aConnection.Open();
                OleDbCommand myCmd = new OleDbCommand(strSql, aConnection);
                i = (int)myCmd.ExecuteScalar();

            }
            catch (Exception ex)
            {
                //MessageBox.Show("ERROR:" + ex.Message);

            }
            finally
            {
                if (aConnection != null)
                    aConnection.Close();

            }

            //当天表为空表,删除表
            if (i <= 0)
            {
                DbOperate del = new DbOperate();

                del.DeleteTable(date1);

            }

            try
            {
                Environment.Exit(Environment.ExitCode);
            }
            catch (Exception)
            {
                //System.Windows.MessageBox.Show("未能正常关闭软件！");
            }

        }

               

        //private void comboBoxTextChanged(object sender, TextChangedEventArgs e)
        //{
        //    if (spOpenSign)
        //    {
        //        if (hbmark)
        //        {
        //            Byte[] temp = new Byte[1];
        //            temp[0] = 0X01;

        //            sp.Write(temp, 0, 1);
        //            hbmark = false;
        //        }
        //    }
        //}

        //private void textboxhb_TextChanged(object sender, TextChangedEventArgs e)
        //{
        //    if (spOpenSign)
        //    {
        //        Byte[] temp = new Byte[6];
        //        temp[0] = 0XE0; temp[1] = 0X00; temp[2] = 0X00;
        //        if (textboxhb.Text.Trim().Length == 0)
        //        {
        //            temp[3] = 0X00;
        //            temp[4] = 0;
        //            temp[5] = 0XE0;
        //            sp.Write(temp, 0, 6);
        //        }
        //        else
        //        {
        //            int rb = Convert.ToInt16(textboxhb.Text.Trim());
        //            temp[3] = (Byte)(rb / 256);
        //            temp[4] = (Byte)(rb % 256);
        //            temp[5] = (Byte)(temp[0] + temp[3] + temp[4]);
        //            sp.Write(temp, 0, 6);
        //        }
        //    }
        //}

        private void textboxhb_KeyUp(object sender, System.Windows.Input.KeyEventArgs e)
        {
            if (spOpenSign)
            {
                Byte[] temp = new Byte[6];
                temp[0] = 0XE0; temp[1] = 0X00; temp[2] = 0X00;
                if (textboxhb.Text.Trim().Length == 0)
                {
                    temp[3] = 0X00;
                    temp[4] = 0;
                    temp[5] = 0XE0;
                    sp.Write(temp, 0, 6);
                }
                else
                {
                    //try
                    //{
                    //    int rb = Convert.ToInt16(textboxhb.Text.Trim());
                    //    temp[3] = (Byte)(rb / 256);
                    //    temp[4] = (Byte)(rb % 256);
                    //    temp[5] = (Byte)(temp[0] + temp[3] + temp[4]);
                    //    sp.Write(temp, 0, 6);
                    //}
                    //catch (Exception)
                    //{

                    //}
                    //finally
                    //{

                    //}

                    //if (string.IsNullOrWhiteSpace(textboxhb.Text.Trim())!=false)
                    //{
                    //    int rb = Convert.ToInt16(textboxhb.Text.Trim());
                    //    temp[3] = (Byte)(rb / 256);
                    //    temp[4] = (Byte)(rb % 256);
                    //    temp[5] = (Byte)(temp[0] + temp[3] + temp[4]);
                    //    sp.Write(temp, 0, 6);
                    //}

                    int tmp;
                    if (int.TryParse(textboxhb.Text.Trim(), out tmp))
                    {
                        //int rb = Convert.ToInt16(textboxhb.Text.Trim());
                        temp[3] = (Byte)(tmp / 256);
                        temp[4] = (Byte)(tmp % 256);
                        temp[5] = (Byte)(temp[0] + temp[3] + temp[4]);
                        sp.Write(temp, 0, 6);
                    }
                    else
                    {

                    }

                }
            }
        }

        private void textboxhb_LostFocus(object sender, RoutedEventArgs e)
        {
            if (spOpenSign)
            {
                Byte[] temp = new Byte[6];
                temp[0] = 0XE0; temp[1] = 0X00; temp[2] = 0X00;
                if (textboxhb.Text.Trim().Length == 0)
                {
                    temp[3] = 0X00;
                    temp[4] = 0;
                    temp[5] = 0XE0;
                    sp.Write(temp, 0, 6);
                }
                else
                {
                    //try
                    //{
                    //    int rb = Convert.ToInt16(textboxhb.Text.Trim());
                    //    temp[3] = (Byte)(rb / 256);
                    //    temp[4] = (Byte)(rb % 256);
                    //    temp[5] = (Byte)(temp[0] + temp[3] + temp[4]);
                    //    sp.Write(temp, 0, 6);
                    //}
                    //catch (Exception)
                    //{

                    //}
                    //finally
                    //{

                    //}

                    //if (string.IsNullOrWhiteSpace(textboxhb.Text.Trim())!=false)
                    //{
                    //    int rb = Convert.ToInt16(textboxhb.Text.Trim());
                    //    temp[3] = (Byte)(rb / 256);
                    //    temp[4] = (Byte)(rb % 256);
                    //    temp[5] = (Byte)(temp[0] + temp[3] + temp[4]);
                    //    sp.Write(temp, 0, 6);
                    //}

                    int tmp;
                    if (int.TryParse(textboxhb.Text.Trim(), out tmp))
                    {
                        //int rb = Convert.ToInt16(textboxhb.Text.Trim());
                        temp[3] = (Byte)(tmp / 256);
                        temp[4] = (Byte)(tmp % 256);
                        temp[5] = (Byte)(temp[0] + temp[3] + temp[4]);
                        sp.Write(temp, 0, 6);
                    }
                    else
                    {

                    }

                }
            }
        }
    }
}
