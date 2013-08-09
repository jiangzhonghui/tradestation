using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.IO;
using System.Net;
using System.Net.Mail;
using System.Net.Mime;
using System.Threading;



namespace SenderSMS_TS8
{
	public partial class Form1 : Form
	{
		private string strWindowsText = "Minder";
        
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

		public Form1()
		{
			InitializeComponent();
		}
		//private string strConnection;

		string phone;
		string command;

		private void Form1_Load(object sender, EventArgs e)
		{
            //string dbfileName = System.Configuration.ConfigurationSettings.AppSettings.Get("DB") + "\\message.mdb";
            //string dbfile = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source={0}";
			//strConnection = string.Format(dbfile, dbfileName);
			phone = System.Configuration.ConfigurationSettings.AppSettings.Get("Phone").Replace(",", "\r\n");
			command = System.Configuration.ConfigurationSettings.AppSettings.Get("SenderMsgCommand");
			string interval = System.Configuration.ConfigurationSettings.AppSettings.Get("TimerInterval");
			this.timer1.Interval = Convert.ToInt32(interval);
		}

		private string PreMessage = string.Empty;
		private IntPtr hWinPreMainWidows;
		private bool CanSend=false;

		private void timer1_Tick(object sender, EventArgs e)
		{
			SendSMSTimerInterval();
		}

		private void SendSmsTimer()
		{	
			SendSMSTimerInterval();
		}

		private void CreateEml(string Message)
		{
			using(System.IO.StreamWriter objEml = new System.IO.StreamWriter(@"C:\Inetpub\mailroot\Drop\"+System.Guid.NewGuid().ToString()+".txt", true,Encoding.UTF8))
			{
				objEml.WriteLine("TradeStation - 电资讯 限价提示");
				objEml.WriteLine("	Order: "+Message);
				objEml.WriteLine("	Account: ");
				objEml.WriteLine("	Occurred: "+System.DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
				objEml.WriteLine("	Signal: ");
				objEml.WriteLine("	Interval: ");
				objEml.WriteLine("	Workspace: ");
				objEml.Close();
			}
		}

		private string strPreID = string.Empty;
        System.IO.StreamWriter objLog =null;


        private void SendSMSTimerInterval()
        {

            this.timer1.Enabled = false;
            objLog = new System.IO.StreamWriter(System.Configuration.ConfigurationSettings.AppSettings.Get("DB") + "\\log_sms_send.txt", true);
            string sql = string.Empty;
            IntPtr hWaitingForSMSSender = IntPtr.Zero;
            IntPtr hWinMainWidowsSMSError = IntPtr.Zero;
            IntPtr hWinMainWidowsSMSError1 = IntPtr.Zero;
            hWaitingForSMSSender = FindWindow("#32770", "SMS Sender");
            if (hWaitingForSMSSender != IntPtr.Zero)
            {
                //sql = "update ta_main set sttime='" + System.DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "',status=2 where status=1 and id=" + strPreID;
                //int i = DBHelper.getInstance(strConnection).ExecuteNonQuery(sql);
                //DBHelper.Close();
                System.Threading.Thread.Sleep(1000);
                PostMessage(hWaitingForSMSSender, WM_CLOSE, 0, "");
                objLog.WriteLine("SMS Sender");
                objLog.WriteLine(sql);
            }

            hWinMainWidowsSMSError = FindWindow("#32770", "SMS Sender 错误");
            hWinMainWidowsSMSError1 = FindWindow("#32770", "SMS Sender Error");
            if (hWinMainWidowsSMSError != IntPtr.Zero || hWinMainWidowsSMSError1 != IntPtr.Zero)
            {
                PostMessage(hWinMainWidowsSMSError, WM_CLOSE, 0, "");
                PostMessage(hWinMainWidowsSMSError1, WM_CLOSE, 0, "");
                //PostMessage(hAllInOneButtonHandle, BM_CLICK, 0, "");
                //sql = "update ta_main set sttime='" + System.DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "',status=0 where status=1 and id=" + strPreID;
                //int i = DBHelper.getInstance(strConnection).ExecuteNonQuery(sql);
                //DBHelper.Close();
                objLog.WriteLine("SMS Sender 错误");
                objLog.WriteLine(sql);
                System.Threading.Thread.Sleep(1000);
            }

            DirectoryInfo dirInfo = new DirectoryInfo(System.Configuration.ConfigurationSettings.AppSettings.Get("Mail"));
            foreach (FileInfo fileInfo in dirInfo.GetFiles())
            {
                String fileName = fileInfo.FullName;

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
                    objLog.WriteLine("");
                    objLog.WriteLine("-----" + System.DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "----");
                    while (line != null)
                    {

                        if (line.StartsWith("TradeStation"))
                        {
                            strTradeStation = line.Replace("TradeStation -", "").Trim() + " ";
                        }
                        if (line.StartsWith("Order:"))
                        { 
                            strOrder = line.Replace("Order:", "").Replace("@ Market","").Trim() +" ";
                        }
                        if (line.StartsWith("Account:"))
                        {
                            strAccount = line.Trim() +" ";
                        }
                        if (line.StartsWith("Filled Price:"))
                        {
                            strSignal = line.Replace("Filled Price:", "Price:").Trim() + " ";
                        }
                        if (line.StartsWith("Duration:"))
                        {
                            strInterval = line.Trim() +" ";
                        }
                         if (line.StartsWith("Current Date:"))
                        {
                            strCurrDate = line.Replace("Current Date:", "").Trim() + " ";
                        }
                        if (strOrder != string.Empty || strCurrDate != string.Empty || strTradeStation != string.Empty || strAccount != string.Empty || strSignal != string.Empty || strWorkspace != string.Empty || strInterval != string.Empty)
                        {
                            objLog.WriteLine(line);
                        }
                        line = file.ReadLine();
                    }

                    file.Close();

                    if (strOrder == string.Empty)
                    {
                        objLog.WriteLine("The file is not Trade Station order notification file.");
                        objLog.WriteLine("----End----");
                    }
                    else
                    {
                        if (System.Configuration.ConfigurationSettings.AppSettings.Get("Send") == "1")
                        {

                            
                            String strMessageContent = strCurrDate + strOrder.Replace("@","") + strSignal;

                            if (System.Configuration.ConfigurationSettings.AppSettings.Get("NotificationWeiXin") == "true")
                            {
                                String from = System.Configuration.ConfigurationSettings.AppSettings.Get("fromMail");
                                String to = System.Configuration.ConfigurationSettings.AppSettings.Get("toMail");
                                String host = System.Configuration.ConfigurationSettings.AppSettings.Get("host");
                                String user = System.Configuration.ConfigurationSettings.AppSettings.Get("username");
                                String pwd = System.Configuration.ConfigurationSettings.AppSettings.Get("password");

                                try
                                {
                                    sendMail(host, user, pwd, from, to, strMessageContent, strMessageContent);
                                }
                                catch (Exception ex)
                                {
                                    objLog.WriteLine("failed to send to weixin..");
                                }
                                
                            }

                            Process myProcess = new Process();
                            myProcess.StartInfo.FileName = command;
                            myProcess.StartInfo.UseShellExecute = false;
                            myProcess.StartInfo.Arguments = " /u /p:" + phone + " /m:\"" + strMessageContent + "\"";

                            objLog.WriteLine(" /u /p:" + phone + " /m:\"" + strMessageContent + "\"");

                            myProcess.StartInfo.WindowStyle = ProcessWindowStyle.Hidden;

                            myProcess.Start();

                            System.Threading.Thread.Sleep(1000);
                            hWinMainWidowsSMSError = FindWindow("#32770", "SMS Sender 错误");
                            hWinMainWidowsSMSError1 = FindWindow("#32770", "SMS Sender Error");
                            if (hWinMainWidowsSMSError != IntPtr.Zero || hWinMainWidowsSMSError1 != IntPtr.Zero)
                            {
                                PostMessage(hWinMainWidowsSMSError, WM_CLOSE, 0, "");
                                PostMessage(hWinMainWidowsSMSError1, WM_CLOSE, 0, "");
                                //PostMessage(hAllInOneButtonHandle, BM_CLICK, 0, "");
                                //sql = "update ta_main set sttime='" + System.DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "',status=0 where status=1 and id=" + strPreID;
                                //int i = DBHelper.getInstance(strConnection).ExecuteNonQuery(sql);
                                //DBHelper.Close();
                                objLog.WriteLine("SMS Sender 错误");
                                //objLog.WriteLine(sql);
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
                                //s//ql = "update ta_main set sttime='" + System.DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "',status=2 where status=1 and id=" + strPreID;
                                //int i = DBHelper.getInstance(strConnection).ExecuteNonQuery(sql);
                                //DBHelper.Close();
                                System.Threading.Thread.Sleep(1000);
                                PostMessage(hWaitingForSMSSender, WM_CLOSE, 0, "");
                                objLog.WriteLine("SMS Sender");
                                //Log_Message(sql);
                            }
                        }
                    }

                    string MailBackUp = System.Configuration.ConfigurationSettings.AppSettings.Get("MailBackUp");
                    if (File.Exists(MailBackUp + "\\" + fileInfo.Name))
                    {
                        File.Delete(MailBackUp + "\\" + fileInfo.Name);
                    }
                    fileInfo.MoveTo(MailBackUp + "\\" + fileInfo.Name);

                }
                catch (Exception ex)
                {
                    objLog.WriteLine(ex.Message.ToString());
                    objLog.WriteLine("----End----");
                }
                finally
                {
                    if (file != null)
                    {
                        file.Close();
                    }
                }
            }
            objLog.Close();
            objLog = null;
            System.Threading.Thread.Sleep(timer1.Interval);
            this.timer1.Enabled = true;
        }

        static bool mailSent = false;
        public static void SendCompletedCallback(object sender, AsyncCompletedEventArgs e)
        {
            // Get the unique identifier for this asynchronous operation.
            String token = (string)e.UserState;

            if (e.Cancelled)
            {
                Console.WriteLine("[{0}] Send canceled.", token);
            }
            if (e.Error != null)
            {
                Console.WriteLine("[{0}] {1}", token, e.Error.ToString());
            }
            else
            {
                Console.WriteLine("Message sent.");
            }
            mailSent = true;
        }

        public static void sendMail(String host, String user,String pwd,String fromMail, String toMail, String subject, String strbody)
        {
            // Command line argument must the the SMTP host.
            SmtpClient client = new SmtpClient(host,25);
            client.Credentials = new System.Net.NetworkCredential(user, pwd);
            // Specify the e-mail sender.
            // Create a mailing address that includes a UTF8 character
            // in the display name.
            MailAddress from = new MailAddress(fromMail, "Tradestation",
            System.Text.Encoding.UTF8);
            // Set destinations for the e-mail message.
            MailAddress to = new MailAddress(toMail);
            // Specify the message content.
            MailMessage message = new MailMessage(from, to);
            message.Body = strbody;
            // Include some non-ASCII characters in body and subject.
            string someArrows = new string(new char[] { '\u2190', '\u2191', '\u2192', '\u2193' });
            message.Body += Environment.NewLine + someArrows;
            message.BodyEncoding = System.Text.Encoding.UTF8;
            message.Subject = subject + someArrows;
            message.SubjectEncoding = System.Text.Encoding.UTF8;
            // Set the method that is called back when the send operation ends.
            client.SendCompleted += new
            SendCompletedEventHandler(SendCompletedCallback);
            // The userState can be any object that allows your callback 
            // method to identify this send operation.
            // For this example, the userToken is a string constant.
            string userState = "test message1";
            client.Send(message);
            //client.SendAsync(message, userState);
            // Clean up.
            message.Dispose();
            Console.WriteLine("Goodbye.");
        }

		private void button1_Click(object sender, EventArgs e)
		{
			this.timer1.Enabled = true;
			this.CanSend = true;
			this.button1.Enabled = false;
			this.button2.Enabled = true;
		}

		private void button2_Click(object sender, EventArgs e)
		{
			this.timer1.Enabled = false;
			this.CanSend = false;
			this.button1.Enabled = true;
			this.button2.Enabled = false;
		}

        private void button3_Click(object sender, EventArgs e)
        {
            String strMessageContent = "this is a testing message";
            if (System.Configuration.ConfigurationSettings.AppSettings.Get("NotificationWeiXin") == "true")
            {
 
                String from = System.Configuration.ConfigurationSettings.AppSettings.Get("fromMail");
                String to = System.Configuration.ConfigurationSettings.AppSettings.Get("toMail");
                String host = System.Configuration.ConfigurationSettings.AppSettings.Get("host");
                String user = System.Configuration.ConfigurationSettings.AppSettings.Get("username");
                String pwd = System.Configuration.ConfigurationSettings.AppSettings.Get("password");

                sendMail(host, user, pwd, from, to, strMessageContent, strMessageContent);

            }
        }
	}
}