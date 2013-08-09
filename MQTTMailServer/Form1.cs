using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Diagnostics;
using System.IO;
using System.Net.Mail;
using System.Net.Sockets;
using System.Net;
using MqttLib;
using System.Runtime.InteropServices;

namespace MQTTMailServer
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            System.Threading.ThreadPool.QueueUserWorkItem(new System.Threading.WaitCallback(StartSMTPServer));
            timer1.Enabled = true;
        }

        public delegate void EventHandler_SetControlText(object sender, string message);
        public delegate void EventHandler_SetControlStatus(object sender);
        public delegate void EventHandler_DataBind(object sender);

        private void Log_Message(string message)
        {
            String log = System.DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + " " + message;
            EventHandler_SetControlText Event_SetText = SetControlText;
            this.txtLog.Invoke(Event_SetText, new object[] { txtLog, log });
            objLog.WriteLine(log);
        }


        private void SetControlText(object control, string message)
        {
            RichTextBox pCard = control as RichTextBox;
            pCard.AppendText(message + "\n");
        }

        TcpListener listener;
        Boolean isStop;

        public void StartSMTPServer(object o)
        {
            
            TcpClient client;
            NetworkStream ns;

            Log_SMTP_Message("listener.Start();...");
            try
            {
                listener.Start();
            }
            catch (Exception ex)
            {
                Log_SMTP_Message(ex.Message.ToString());
                return;
            }
            
            isStop = false;
            Log_SMTP_Message("log_smtp..");
            objLogSMTP = new System.IO.StreamWriter(Application.StartupPath + "\\log_smtp.txt", true);

            Log_SMTP_Message("Awaiting connection...");

            while (isStop == false)
            {
                client = listener.AcceptTcpClient();

                Log_SMTP_Message("Connection accepted!");

                ns = client.GetStream();

                try
                {
                    using (StreamWriter writer = new StreamWriter(ns))
                    {
                        writer.WriteLine("220 localhost SMTP server ready.");
                        writer.Flush();

                        using (StreamReader reader = new StreamReader(ns))
                        {
                            string response = reader.ReadLine();
                            Log_SMTP_Message(response);
                            if (!response.StartsWith("HELO") && !response.StartsWith("EHLO"))
                            {
                                writer.WriteLine("500 UNKNOWN COMMAND");
                                writer.Flush();
                                ns.Close();
                                client.Close();
                                continue;
                            }

                            string remote = response.Replace("HELO", string.Empty).Replace("EHLO", string.Empty).Trim();

                            writer.WriteLine("250 localhost Hello " + remote);
                            writer.Flush();

                            response = reader.ReadLine();
                            Log_SMTP_Message(response);
                            if (!response.StartsWith("MAIL FROM:"))
                            {
                                writer.WriteLine("500 UNKNOWN COMMAND");
                                writer.Flush();
                                ns.Close();
                                client.Close();
                                continue;
                            }

                            remote = response.Replace("RCPT TO:", string.Empty).Trim();
                            writer.WriteLine("250 " + remote + " I like that guy too!");
                            writer.Flush();

                            response = reader.ReadLine();
                            Log_SMTP_Message(response);
                            if (!response.StartsWith("RCPT TO:"))
                            {
                                writer.WriteLine("500 UNKNOWN COMMAND");
                                writer.Flush();
                                ns.Close();
                                client.Close();
                                continue;
                            }

                            remote = response.Replace("MAIL FROM:", string.Empty).Trim();
                            writer.WriteLine("250 " + remote + " I like that guy!");
                            writer.Flush();

                            response = reader.ReadLine();
                            Log_SMTP_Message(response);
                            if (response.Trim() != "DATA")
                            {
                                writer.WriteLine("500 UNKNOWN COMMAND");
                                writer.Flush();
                                ns.Close();
                                client.Close();
                                return;
                            }

                            writer.WriteLine("354 Enter message. When finished, enter \".\" on a line by itself");
                            writer.Flush();

                            int counter = 0;
                            StringBuilder message = new StringBuilder();
                            Log_SMTP_Message("---Start while---");
                            //System.IO.StreamWriter objMail = new System.IO.StreamWriter(System.Configuration.ConfigurationSettings.AppSettings.Get("DB") + "\\" + System.DateTime.Now.ToString("yyyyMMddHHmmss") + System.Guid.NewGuid().ToString() + ".txt", true);
                            while ((response = reader.ReadLine().Trim()) != ".")
                            {

                     
//FROM:ts@lou9593e2a556e
//TO:<administrator@lou9593e2a556e>
//SUBJECT:TradeStation - Order has been filled for CLQ12

//TradeStation - Order has been filled for CLQ12
//Order: Buy 1 CLQ12 @ Market
//Qty Filled: 1
//Filled Price: 85.9300
//Duration: Day
//Route: N/A
//Account: SIM569903F
//Order#: 2-4350-1621
                                if (response.StartsWith("TradeStation") || response.StartsWith("Order") || response.StartsWith("Filled Price") || response.StartsWith("Duration") || response.StartsWith("Account") || response.StartsWith("This is a test message"))
                                {
                                    response=response.Replace("TradeStation -","").Replace("Order: ","");
                                    Log_SMTP_Message(response);

                                    message.Append(response+",");
                                }
                                
                                counter++;

                                if (counter == 1000000)
                                {
                                    ns.Close();
                                    client.Close();
                                    return;  // Seriously? 1 million lines in a message?
                                }
                            }
                            //objMail.Close();
                            Log_SMTP_Message("---End while---");
                            writer.WriteLine("250 OK");
                            writer.Flush();
                            ns.Close();
                            // Insert "message" into DB
                            Log_SMTP_Message("Received message:");
                            Log_SMTP_Message(message.ToString());
                            SendMessage(message.ToString(), "20120512");

                        }
                    }
                }
                catch (Exception e)
                {
                    Log_Message(e.Message.ToString());
                }
                finally
                {
                    client.Close();
                }
                
            }

        }

        private void SendMessage(String strOrder,String strCurrDate)
        {
            Log_Message("Start send message to MQTT server");
            //mqtt_server
            string connString = System.Configuration.ConfigurationSettings.AppSettings.Get("mqtt_server");
            IMqtt client = MqttLib.MqttClientFactory.CreateClient(connString, "tsorderservice");
            client.Connect();
            String mqttString = "mqtt_send_sms|" + strOrder + "";

            Log_Message(mqttString);

            if (client != null && client.IsConnected)
            {
                try
                {
                    string topic = "mqtt/download";
                    int iResult=client.Publish(
                            topic,
                            new MqttPayload(mqttString),
                            QoS.BestEfforts,
                            false
                        );
                    Log_Message("Send message result :" + iResult);
                }
                catch(Exception ex) 
                {
                    Log_Message("Failed to send message result :" + ex.Message.ToString());
                }
                finally
                {
                    client.Disconnect();
                }
            }
            //try
            //{
            //    string dbfileName = System.Configuration.ConfigurationSettings.AppSettings.Get("DB") + "\\message.mdb";
            //    string dbfile = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source={0}";
            //    string strConnection = string.Format(dbfile, dbfileName);
            //    string sql = "insert into ta_main(Msg,crtime) values('" + strOrder + "','" + strCurrDate + "')";
            //    int j = DBHelper.getInstance(strConnection).ExecuteNonQuery(sql);
            //    DBHelper.Close();
            //    Log_Message(sql);
            //}
            //catch (Exception ex)
            //{
            //    Log_Message(ex.Message.ToString());
            //}
           
            Log_Message("----End----");
        }

        private System.IO.StreamWriter objLog = null;

       
        private System.IO.StreamWriter objLogSMTP = null;

        public void Log_SMTP_Message(string logmessage)
        {
            if (objLogSMTP != null)
            {
                objLogSMTP.WriteLine(logmessage);
            }
            Log_Message(logmessage);
        }

        protected bool LoadMail(string fileName)
        {
            bool IsLoadMail = false;
            StreamReader file = null;
            try
            {
                string strCurrDate = string.Empty;
                string strOrder = string.Empty;
                string strAccount = string.Empty;
                string strSignal = string.Empty;
                string strInterval = string.Empty;
                string strWorkspace = string.Empty;
                string strTradeStation = string.Empty;

                file = new StreamReader(fileName, Encoding.Default);
                string line = file.ReadLine();
                Log_Message("");
                Log_Message("-----" + System.DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "----");
                while (line != null)
                {

                    if (line.StartsWith("TradeStation"))
                    {
                        strTradeStation = line.Replace("TradeStation -", "").Trim();
                    }
                    if (line.StartsWith("	Order:"))
                    {
                        strOrder = line.Replace("	Order:", "").Trim();
                    }
                    if (line.StartsWith("	Account:"))
                    {
                        strAccount = line.Replace("Account:", "").Trim();
                    }
                    if (line.StartsWith("	Occurred:"))
                    {
                        strCurrDate = line.Replace("	Occurred:", "").Trim();
                    }
                    if (line.StartsWith("	Signal:"))
                    {
                        strSignal = line.Replace("	Signal:", "").Trim();
                    }
                    if (line.StartsWith("	Interval:"))
                    {
                        strInterval = line.Replace("	Interval:", "").Trim();
                    }

                    if (line.StartsWith("	Workspace:"))
                    {
                        strWorkspace = line.Replace("	Workspace:", "").Trim();
                    }
                    if (strOrder != string.Empty || strCurrDate != string.Empty || strTradeStation != string.Empty || strAccount != string.Empty || strSignal != string.Empty || strWorkspace != string.Empty || strInterval != string.Empty)
                    {
                        Log_Message(line);
                    }
                    line = file.ReadLine();
                }

                file.Close();

                if (strOrder == string.Empty || strCurrDate == string.Empty)
                {
                    Log_Message("The file is not Trade Station order notification file.");
                    Log_Message("----End----");
                }
                else
                {
                    //mqtt_server
                    string connString = System.Configuration.ConfigurationSettings.AppSettings.Get("mqtt_server");
                    IMqtt client = MqttLib.MqttClientFactory.CreateClient(connString, "tsorderservice");
                    string phoneNumber = System.Configuration.ConfigurationSettings.AppSettings.Get("Phone");
                    client.Connect();
                    String mqttString = "mqtt_send_sms|" + strOrder + "";

                    if (client != null && client.IsConnected)
                    {
                        try
                        {
                            string topic = "mqtt/download";
                            client.Publish(
                                    topic,
                                    new MqttPayload(strOrder),
                                    QoS.BestEfforts,
                                    true
                                );
                        }
                        catch { }
                        finally
                        {
                            client.Disconnect();
                        }
                    }
                    string dbfileName = Application.StartupPath + "\\message.mdb";
                    string dbfile = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source={0}";
                    string strConnection = string.Format(dbfile, dbfileName);
                    string sql = "insert into ta_main(Msg,crtime) values('" + strOrder + "','" + strCurrDate + "')";
                    int j = DBHelper.getInstance(strConnection).ExecuteNonQuery(sql);
                    DBHelper.Close();
                    Log_Message(sql);
                    Log_Message("----End----");

                }

                FileInfo fileInfo = new FileInfo(fileName);
                string MailBackUp = System.Configuration.ConfigurationSettings.AppSettings.Get("MailBackUp");
                if (File.Exists(MailBackUp + "\\" + fileInfo.Name))
                {
                    File.Delete(MailBackUp + "\\" + fileInfo.Name);
                }
                fileInfo.MoveTo(MailBackUp + "\\" + fileInfo.Name);

                IsLoadMail = true;
            }
            catch (Exception ex)
            {
                Log_Message(ex.Message.ToString());
                Log_Message("----End----");
                IsLoadMail = false;
            }
            finally
            {
                if (file != null)
                {
                    file.Close();
                }
            }
            return IsLoadMail;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            timer1.Enabled = false;
            isStop = false;
            listener.Stop();

        }

        private void Form1_Load(object sender, EventArgs e)
        {
            string dbfileName = Application.StartupPath+ "\\winquote.mdb";
            string dbfile = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source={0}";
            strConnection = string.Format(dbfile, dbfileName);

            listener = new TcpListener(IPAddress.Loopback, 25);
           
            objLog = new System.IO.StreamWriter(Application.StartupPath + "\\log.txt", true);
			
        }
        private string strConnection;
        string phone;
        string command;

        private void SendSMSTimerInterval()
        {

            this.timer1.Enabled = false;
            string sql = string.Empty;
            IntPtr hWaitingForSMSSender = IntPtr.Zero;
            IntPtr hWinMainWidowsSMSError = IntPtr.Zero;
            IntPtr hWinMainWidowsSMSError1 = IntPtr.Zero;
            hWaitingForSMSSender = FindWindow("#32770", "SMS Sender");
            if (hWaitingForSMSSender != IntPtr.Zero)
            {
                sql = "update ta_main set sttime='" + System.DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "',status=2 where status=1 and id=" + strPreID;
                int i = DBHelper.getInstance(strConnection).ExecuteNonQuery(sql);
                //DBHelper.Close();
                System.Threading.Thread.Sleep(1000);
                PostMessage(hWaitingForSMSSender, WM_CLOSE, 0, "");
                Log_Message("SMS Sender");
                Log_Message(sql);
            }

            hWinMainWidowsSMSError = FindWindow("#32770", "SMS Sender ´íÎó");
            hWinMainWidowsSMSError1 = FindWindow("#32770", "SMS Sender Error");
            if (hWinMainWidowsSMSError != IntPtr.Zero || hWinMainWidowsSMSError1 != IntPtr.Zero)
            {
                PostMessage(hWinMainWidowsSMSError, WM_CLOSE, 0, "");
                PostMessage(hWinMainWidowsSMSError1, WM_CLOSE, 0, "");
                //PostMessage(hAllInOneButtonHandle, BM_CLICK, 0, "");
                sql = "update ta_main set sttime='" + System.DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "',status=0 where status=1 and id=" + strPreID;

                int i = DBHelper.getInstance(strConnection).ExecuteNonQuery(sql);
                //DBHelper.Close();
                Log_Message("SMS Sender ´íÎó");
                Log_Message(sql);
                System.Threading.Thread.Sleep(1000);
            }

            sql = "select * from ta_main where status = 0 order by id";
            DataSet dsMessage = DBHelper.getInstance(strConnection).ExecuteDataset(sql);
            if (dsMessage.Tables[0].Rows.Count > 0)
            {
                DataRow drw = dsMessage.Tables[0].Rows[0];

                string strMessageContents = drw[1].ToString();
                string strId = drw[0].ToString();
                Log_Message(strMessageContents);

                sql = "update ta_main set status=1 where id=" + strId;
                DBHelper.getInstance(strConnection).ExecuteScalar(sql);
                strPreID = strId;

                Log_Message(" /u /p:" + phone + " /m:\"" + strMessageContents + "\"");

                //DBHelper.Close();

                if (System.Configuration.ConfigurationSettings.AppSettings.Get("Send") == "1")
                {
                    Process myProcess = new Process();
                    myProcess.StartInfo.FileName = command;
                    myProcess.StartInfo.UseShellExecute = false;
                    myProcess.StartInfo.Arguments = " /u /p:" + phone + " /m:\"" + strMessageContents + "\"";

                    myProcess.StartInfo.WindowStyle = ProcessWindowStyle.Hidden;

                    myProcess.Start();

                    System.Threading.Thread.Sleep(1000);
                    hWinMainWidowsSMSError = FindWindow("#32770", "SMS Sender ´íÎó");
                    hWinMainWidowsSMSError1 = FindWindow("#32770", "SMS Sender Error");
                    if (hWinMainWidowsSMSError != IntPtr.Zero || hWinMainWidowsSMSError1 != IntPtr.Zero)
                    {
                        PostMessage(hWinMainWidowsSMSError, WM_CLOSE, 0, "");
                        PostMessage(hWinMainWidowsSMSError1, WM_CLOSE, 0, "");
                        //PostMessage(hAllInOneButtonHandle, BM_CLICK, 0, "");
                        sql = "update ta_main set sttime='" + System.DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "',status=0 where status=1 and id=" + strPreID;

                        int i = DBHelper.getInstance(strConnection).ExecuteNonQuery(sql);
                        //DBHelper.Close();
                        this.Log_Message("SMS Sender ´íÎó");
                        Log_Message(sql);
                        System.Threading.Thread.Sleep(1000);
                    }
                    else
                    {
                        myProcess.WaitForExit(4000);
                    }
                    myProcess.Close();
                    hWaitingForSMSSender = FindWindow("#32770", "SMS Sender");
                    if (hWaitingForSMSSender != IntPtr.Zero)
                    {
                        sql = "update ta_main set sttime='" + System.DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "',status=2 where status=1 and id=" + strPreID;
                        int i = DBHelper.getInstance(strConnection).ExecuteNonQuery(sql);
                        //DBHelper.Close();
                        System.Threading.Thread.Sleep(1000);
                        PostMessage(hWaitingForSMSSender, WM_CLOSE, 0, "");
                        Log_Message("SMS Sender");
                        Log_Message(sql);
                    }
                }
                sql = "update ta_main set sttime='" + System.DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "',status=2 where id=" + strPreID;
                DBHelper.getInstance(strConnection).ExecuteNonQuery(sql);
                //DBHelper.Close();
                Log_Message(sql);
            }
            System.Threading.Thread.Sleep(timer1.Interval);
            this.timer1.Enabled = true;
        }

        private string strPreID = string.Empty;

        private void timer1_Tick(object sender, EventArgs e)
        {
            SendSMSTimerInterval();
        }

        int WM_SETFOCUS = 0x0007;
        int BM_CLICK = 0x00F5;
        int WM_SETTEXT = 0x000C;
        int CB_SELECTSTRING = 0x014D;
        int CB_INSERTSTRING = 0x014A;
        int CB_ADDSTRING = 0x0143;
        int WM_KEYDOWN = 0x0100;
        int VK_RETURN = 0x0D;
        int WM_KEYUP = 0x0101;
        int WM_CHAR = 0x0102;
        int WM_CLOSE = 0x0010;
        int CB_SETCURSEL = 0x014E;





        [DllImport("user32.dll")]
        private static extern void BlockInput(bool Block);

        [DllImport("user32.dll")]
        private static extern IntPtr GetDesktopWindow();

        [DllImport("gdi32.dll", CharSet = CharSet.Auto)]
        private static extern bool TextOut(IntPtr hdc, int nXStart, int nYStart, string lpString, int cbString);

        [DllImport("user32.dll", EntryPoint = "FindWindow", SetLastError = true)]
        private static extern IntPtr FindWindow(string lpClassName, string lpWindowName);

        [DllImport("user32.dll", EntryPoint = "FindWindowEx", SetLastError = true)]
        private static extern IntPtr FindWindowEx(IntPtr hwndParent, IntPtr hwndChildAfter, string lpszClass, string lpszWindow);

        [DllImport("user32.dll", EntryPoint = "SetForegroundWindow", SetLastError = true)]
        private static extern void SetForegroundWindow(IntPtr hwnd);

        [DllImport("user32.dll", SetLastError = true)]
        public static extern bool EnumWindows(IntPtr lpEnumFunc, uint lParam);
        [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)]
        public static extern int GetWindowTextLength(IntPtr hWnd);
        [DllImport("user32.dll")]
        public static extern int GetWindowText(IntPtr hWnd, [Out] StringBuilder lpString, int nMaxCount);

        public delegate int EnumWindowsProc(IntPtr hwnd, IntPtr lParam);

        public delegate bool EnumChildWindowsProc(IntPtr hwnd, uint lParam);

        [DllImport("user32.dll", EntryPoint = "EnumChildWindows")]
        public static extern bool EnumChildWindows(IntPtr hWndParent, EnumChildWindowsProc lpEnumFunc, int lParam);

        [DllImport("user32.dll", EntryPoint = "SwitchToThisWindow")]
        public static extern bool SwitchToThisWindow(IntPtr hWndParent, bool lParam);



        [DllImport("user32.dll", EntryPoint = "GetClassName")]
        public static extern int GetClassName(IntPtr hwnd, StringBuilder lpClassName, int nMaxCount);


        [DllImport("user32.dll", EntryPoint = "SendMessageA")]
        public static extern int SendMessage(IntPtr hwnd, int wMsg, int wParam, String lParam);

        [DllImport("user32.dll", EntryPoint = "PostMessageA")]
        public static extern int PostMessage(IntPtr hwnd, int wMsg, int wParam, String lParam);

        [DllImport("user32.dll", EntryPoint = "PostMessageA")]
        public static extern int PostMessage(IntPtr hwnd, int wMsg, int wParam, int lParam);

    }
}