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
		private string strConnection;
		string phone;
		string command;
		string[] activeTime;

		private void Form1_Load(object sender, EventArgs e)
		{
			string dbfileName = System.Configuration.ConfigurationSettings.AppSettings.Get("DB") + "\\winquote.mdb";
			string dbfile = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source={0}";
			strConnection = string.Format(dbfile, dbfileName);
			string interval = System.Configuration.ConfigurationSettings.AppSettings.Get("TimerInterval");
			char[] sp ={ ',' };
			activeTime = System.Configuration.ConfigurationSettings.AppSettings.Get("ActiveTime").Split(sp);
			this.timer1.Interval = Convert.ToInt32(interval);
		}

		private string PreMessage = string.Empty;
		private IntPtr hWinPreMainWidows;
		private bool CanSend=false;

		private void FindMinder()
		{
			try
			{
				this.timer1.Enabled = false;
				IntPtr hWinMainWidows=IntPtr.Zero;
				hWinMainWidows = FindWindow(null, "Minder");
				if (hWinMainWidows != IntPtr.Zero)
				{
					IntPtr hWinMessageTextBox;
					IntPtr hAllInOneButtonHandle;
					IntPtr hHandle1 = FindWindowEx(hWinMainWidows, IntPtr.Zero, "Static", null);
					IntPtr hHandle2 = FindWindowEx(hWinMainWidows, hHandle1, "Static", null);
					IntPtr hHandle3 = FindWindowEx(hWinMainWidows, hHandle2, "Button", null);
					hWinMessageTextBox = hHandle2;
					hAllInOneButtonHandle = hHandle3;
					StringBuilder strWarningMessage = new StringBuilder(50);
					GetWindowText(hWinMessageTextBox, strWarningMessage, 50);
					//strWarningMessage.Append("最小限额:HSINF1 (21675) o21740");
					if (strWarningMessage.Length > 0)
					{
						string sql = "select count(*) from ta_main where Msg='" + strWarningMessage.ToString() + "'";
						int j = (int)DBHelper.getInstance(strConnection).ExecuteScalar(sql);

						if (j > 0)
						{
							return;
						}

						hWinPreMainWidows = hWinMainWidows;
						PreMessage = strWarningMessage.ToString();
						System.IO.StreamWriter objLog = new System.IO.StreamWriter(System.Configuration.ConfigurationSettings.AppSettings.Get("DB") + "\\smslog.txt", true);
						try
						{
							//string strCurrDate = System.DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
							//string sql = "insert into ta_main(Msg,crtime) values('" + strWarningMessage + "','" + strCurrDate + "')";
							//int j = DBHelper.getInstance(strConnection).ExecuteNonQuery(sql);
							//DBHelper.Close();
							//objLog.WriteLine(sql);

							string regExpress = strWarningMessage.ToString();// "最大限额:388 (130) 130.005(A|B)";
							regExpress = regExpress.Replace("A", "").Replace("B", "").Replace(" ", "").Trim();

							objLog.WriteLine(regExpress);

							bool IsMaxValue = false;
							bool IsAddSymbol = false;
							bool IsSendAlert = false;
							if (regExpress.StartsWith("最大限额:"))
							{
								IsMaxValue = true;
							}
							else if (regExpress.StartsWith("最小限额:"))
							{
								IsMaxValue = false;
							}
							regExpress = regExpress.Replace("最大限额:", "").Replace("最小限额:", "");
							int iStartPos = regExpress.IndexOf("(");
							int iEndPos = regExpress.IndexOf(")");
							string strSymbol = regExpress.Substring(0, iStartPos);
							string strWValue = regExpress.Substring(iStartPos + 1, iEndPos - iStartPos - 1);

							if (!regExpress.Substring(iEndPos + 1).Contains("o"))
							{
								
								string strCurrValue = regExpress.Substring(iEndPos + 1);
								try
								{
									double  TestValue = Convert.ToDouble(strCurrValue);
								}
								catch (Exception ex)
								{
									throw ex;
								}
								
								objLog.WriteLine(strSymbol);
								objLog.WriteLine(strWValue);
								objLog.WriteLine(strCurrValue);

								decimal dCurrMaxValue, dCurrMinValue;
								int iCurrMaxCounter, iCurrMinCounter;
								sql = "select maxprice,minprice,maxcounter,mincounter from ta_main where symbol='" + strSymbol + "'";
								objLog.WriteLine(sql);
								DataSet ds = DBHelper.getInstance(strConnection).ExecuteDataset(sql);
								if (ds.Tables[0].Rows.Count <= 0)
								{
									dCurrMaxValue = 0;
									dCurrMinValue = 0;
									iCurrMaxCounter = 0;
									iCurrMinCounter = 0;
									IsAddSymbol = true;
								}
								else
								{
									dCurrMaxValue = (decimal)ds.Tables[0].Rows[0]["maxprice"];
									dCurrMinValue = (decimal)ds.Tables[0].Rows[0]["minprice"];
									iCurrMaxCounter = (int)ds.Tables[0].Rows[0]["maxcounter"];
									iCurrMinCounter = (int)ds.Tables[0].Rows[0]["mincounter"];
									IsAddSymbol = false;
								}

								if (!IsAddSymbol)
								{
									if (IsMaxValue)
									{
										if (Convert.ToDecimal(strWValue) != dCurrMaxValue)
										{
											IsSendAlert = true;
											iCurrMaxCounter = 0;
										}
										else
										{
											iCurrMaxCounter = iCurrMaxCounter + 1;
										}
										sql = "update ta_main set maxprice=" + strWValue + ", currprice=" + strCurrValue + ",Msg='" + strWarningMessage.ToString() + "',crtime='" + System.DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "' where symbol='" + strSymbol + "'";
									}
									else
									{
										if (Convert.ToDecimal(strWValue) != dCurrMinValue)
										{
											IsSendAlert = true;
											iCurrMinCounter = 1;
										}
										else
										{
											iCurrMinCounter = iCurrMinCounter + 1;
										}
										sql = "update ta_main set minprice=" + strWValue + ", currprice=" + strCurrValue + ",Msg='" + strWarningMessage.ToString() + "',crtime='" + System.DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "' where symbol='" + strSymbol + "'";
									}
								}
								else
								{
									if (IsMaxValue)
									{
										sql = "insert into ta_main(maxprice,currprice,symbol,maxcounter,Msg,crtime) values(" + strWValue + ", " + strCurrValue + ",'" + strSymbol + "',1,'" + strWarningMessage.ToString() + "','" + System.DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "')";
									}
									else
									{
										sql = "insert into ta_main(minprice,currprice,symbol,mincounter,Msg,crtime) values(" + strWValue + ", " + strCurrValue + ",'" + strSymbol + "',1,'" + strWarningMessage.ToString() + "','" + System.DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "')";
									}
									IsSendAlert = true;
								}

								if (IsSendAlert)
								{
									objLog.WriteLine("CreateEml:" + strWarningMessage.ToString());
									CreateEml(strWarningMessage.ToString());
								}
								//SwitchToThisWindow(hWinMainWidows, false);
								//System.Threading.Thread.Sleep(1000);
								PostMessage(hAllInOneButtonHandle, WM_SETFOCUS, 0, "");
								PostMessage(hAllInOneButtonHandle, BM_CLICK, 0, "");
								//System.Threading.Thread.Sleep(2000);
								objLog.WriteLine(sql);
								//SendKeys.SendWait("{ENTER}");
								objLog.WriteLine("Close");

								//sql = "insert into ta_main(Msg,crtime) values('" + strWarningMessage.ToString() + "','" + System.DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "')";
								j = DBHelper.getInstance(strConnection).ExecuteNonQuery(sql);
							}
						}
						catch (Exception ex)
						{
							objLog.WriteLine(ex.Message.ToString());
						}
						finally
						{
							objLog.Close();
						}
					}
				}
				System.Threading.Thread.Sleep(this.timer1.Interval);
			}
			catch (Exception ex)
			{
				System.IO.StreamWriter objLog = new System.IO.StreamWriter(System.Configuration.ConfigurationSettings.AppSettings.Get("DB") + "\\smslog1.txt", true);
				objLog.WriteLine(ex.Message.ToString());
				objLog.Close();
				objLog = null;
			}
			finally
			{
				this.timer1.Enabled = true;
			}
		}

		private void timer1_Tick(object sender, EventArgs e)
		{
			int currTime=Convert.ToInt32(DateTime.Now.ToString("HHmm"));
			if ((currTime>Convert.ToInt32(activeTime[0]) && currTime<=Convert.ToInt32(activeTime[1])) || (currTime>Convert.ToInt32(activeTime[2]) && currTime<=Convert.ToInt32(activeTime[3])))
			{
				FindMinder();
			}
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

		private void SendSMSTimerInterval()
		{

			this.timer1.Enabled = false;
			System.IO.StreamWriter objLog = new System.IO.StreamWriter(System.Configuration.ConfigurationSettings.AppSettings.Get("DB") + "\\smslog.txt", true);
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
				sql = "update ta_main set sttime='" + System.DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "',status=0 where status=1 and id=" + strPreID;

				int i = DBHelper.getInstance(strConnection).ExecuteNonQuery(sql);
				//DBHelper.Close();
				objLog.WriteLine("SMS Sender 错误");
				objLog.WriteLine(sql);
				System.Threading.Thread.Sleep(1000);
			}

			sql = "select * from ta_main where status = 0 order by id";
			DataSet dsMessage = DBHelper.getInstance(strConnection).ExecuteDataset(sql);
			if (dsMessage.Tables[0].Rows.Count > 0)
			{
				DataRow drw = dsMessage.Tables[0].Rows[0];

				string strMessageContents = drw[1].ToString();
				string strId = drw[0].ToString();
				objLog.WriteLine(strMessageContents);

				sql = "update ta_main set status=1 where id=" + strId;
				DBHelper.getInstance(strConnection).ExecuteScalar(sql);
				strPreID = strId;

				objLog.WriteLine(" /u /p:" + phone + " /m:\"" + strMessageContents + "\"");

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
					hWinMainWidowsSMSError = FindWindow("#32770", "SMS Sender 错误");
					hWinMainWidowsSMSError1 = FindWindow("#32770", "SMS Sender Error");
					if (hWinMainWidowsSMSError != IntPtr.Zero || hWinMainWidowsSMSError1 != IntPtr.Zero)
					{
						PostMessage(hWinMainWidowsSMSError, WM_CLOSE, 0, "");
						PostMessage(hWinMainWidowsSMSError1, WM_CLOSE, 0, "");
						//PostMessage(hAllInOneButtonHandle, BM_CLICK, 0, "");
						sql = "update ta_main set sttime='" + System.DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "',status=0 where status=1 and id=" + strPreID;

						int i = DBHelper.getInstance(strConnection).ExecuteNonQuery(sql);
						//DBHelper.Close();
						objLog.WriteLine("SMS Sender 错误");
						objLog.WriteLine(sql);
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
						objLog.WriteLine("SMS Sender");
						objLog.WriteLine(sql);
					}
				}
				sql = "update ta_main set sttime='" + System.DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "',status=2 where id=" + strPreID;
				DBHelper.getInstance(strConnection).ExecuteNonQuery(sql);
				//DBHelper.Close();
				objLog.WriteLine(sql);
			}


			objLog.Close();
			objLog = null;
			System.Threading.Thread.Sleep(timer1.Interval);
			this.timer1.Enabled = true;
		}

		private void SendSMS()
		{
			
		}

		private void CheckFailedSMS(string strID)
		{
			IntPtr hWinMainWidowsSMS;
			hWinMainWidowsSMS = FindWindow("#32770", "SMS Sender 错误");
			if (hWinMainWidowsSMS != IntPtr.Zero)
			{
				PostMessage(hWinMainWidowsSMS, WM_CLOSE, 0, "");
				//PostMessage(hAllInOneButtonHandle, BM_CLICK, 0, "");
				DBHelper.getInstance(strConnection).ExecuteScalar("update ta_main set status=2 where id=" + strID);
				DBHelper.Close();
			}
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
            String strMessageContents = "bbbbbbbbbbbbbb,this is tesing.";
            String phone = "62760119";

            Process myProcess = new Process();
            myProcess.StartInfo.FileName = command;
            myProcess.StartInfo.UseShellExecute = false;
            myProcess.StartInfo.Arguments = " /u /p:" + phone + " /m:\"" + strMessageContents + "\"";

            myProcess.StartInfo.WindowStyle = ProcessWindowStyle.Hidden;

            myProcess.Start();


            //string connString = System.Configuration.ConfigurationSettings.AppSettings.Get("mqtt_server");
            //MqttLib.IMqtt client = MqttLib.MqttClientFactory.CreateClient(connString, "tsorderservice");

            //client.Connect();
            //String strOrder = "mqtt_send_sms|bbbbbbbbbbbbbb";

            //if (client != null && client.IsConnected)
            //{
            //    try
            //    {
            //        string topic = "mqtt/download";
            //        int iResult = client.Publish(
            //                topic,
            //                new MqttLib.MqttPayload(strOrder),
            //                MqttLib.QoS.BestEfforts,
            //                false
            //            );
            //    }
            //    catch (Exception ex)
            //    {
            //        throw ex;
            //    }
            //    finally
            //    {
            //        client.Disconnect();
            //    }
            //}
        }
	}
}