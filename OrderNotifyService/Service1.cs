using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.ServiceProcess;
using System.Text;
using System.IO;
using System.Net.Mail;
using System.Net.Sockets;
using System.Net;
using System.Net.Mime;
using System.Threading;

namespace OrderNotifyService
{
	public partial class Service1 : ServiceBase
	{
		public Service1()
		{
			InitializeComponent();
		}
		private bool isStop = false;

		protected override void OnStart(string[] args)
		{
			// TODO: Add code here to start your service.
			System.Threading.ThreadPool.QueueUserWorkItem(new System.Threading.WaitCallback(StartSMTPServer), null);

			//DirectoryInfo dinfo=new DirectoryInfo(@"C:\Inetpub\mailroot\Drop");
			
            //objLog = new System.IO.StreamWriter(System.Configuration.ConfigurationSettings.AppSettings.Get("DB") + "\\log.txt", true);
			
            //FileInfo[] fileInfoCol = dinfo.GetFiles();
            //int iFileCounter = fileInfoCol.Length;
            //int i=0;
            //while (i < iFileCounter)
            //{
            //    string fileName = fileInfoCol[i].FullName;
            //    bool IsLoadMail=LoadMail(fileName);
            //    if (IsLoadMail)
            //    {
            //        i++;
            //    }
            //    else
            //    {
            //        System.Threading.Thread.Sleep(1000);
            //        int j = 0;
            //        IsLoadMail = LoadMail(fileName);
            //        while (j < 5 && !IsLoadMail)
            //        {
            //            j++;
            //            System.Threading.Thread.Sleep(1000);
            //            IsLoadMail = LoadMail(fileName);
            //        }
            //        i++;
            //    }
            //}
			

            //objLog.Close();
            //objLog = null;

            //this.fileSystemWatcher1.EnableRaisingEvents = true;

            //System.Threading.ThreadPool.QueueUserWorkItem(new System.Threading.WaitCallback(StartSMTPServer));

		}

		protected override void OnStop()
		{
			// TODO: Add code here to perform any tear-down necessary to stop your service.
			//this.fileSystemWatcher1.EnableRaisingEvents = false;
			isStop = true;
		}

		protected override void OnPause()
		{
			//this.fileSystemWatcher1.EnableRaisingEvents = false;
			isStop = true;
		}

		protected override void OnShutdown()
		{
			//this.fileSystemWatcher1.EnableRaisingEvents = false;
			isStop = true;
		}

		protected override void OnContinue()
		{
			//this.fileSystemWatcher1.EnableRaisingEvents = true;
			isStop = false;
		}

		public void StartSMTPServer(object o)
		{
			TcpListener listener = new TcpListener(IPAddress.Loopback, 25);
			TcpClient client;
			NetworkStream ns;
            objLogSMTP = new System.IO.StreamWriter(System.Configuration.ConfigurationSettings.AppSettings.Get("DB") + "\\log_smtp.txt", true);
           
            try
            {
                listener.Start();
                Log_SMTP_Message("listener.Start();...");
            }
            catch (Exception ex)
            {
                Log_SMTP_Message(ex.Message.ToString());
                objLogSMTP.Close();
                return;
            }
			
			Log_SMTP_Message("Awaiting connection...");

			while (isStop==false)
			{
				client = listener.AcceptTcpClient();
				Log_SMTP_Message("Connection accepted!");

				ns = client.GetStream();

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
							return;
						}

						writer.WriteLine("354 Enter message. When finished, enter \".\" on a line by itself");
						writer.Flush();

						int counter = 0;
						StringBuilder message = new StringBuilder();
						Log_SMTP_Message("---Start while---");
                        System.IO.StreamWriter objMail = new System.IO.StreamWriter(System.Configuration.ConfigurationSettings.AppSettings.Get("Mail") + "\\" + System.DateTime.Now.ToString("yyyyMMddHHmmss") + System.Guid.NewGuid().ToString() + ".txt", true);

                        while ((response = reader.ReadLine().Trim()) != ".")
						{
							Log_SMTP_Message(response);
							objMail.WriteLine(response);
							message.AppendLine(response);
							counter++;

							if (counter == 1000)
							{
								ns.Close();
								return;  // Seriously? 1 million lines in a message?
							}
						}
                        objMail.WriteLine("Current Date:" + System.DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
						objMail.Close();
						Log_SMTP_Message("---End while---");
						writer.WriteLine("250 OK");
						writer.Flush();
						ns.Close();

                        if (System.Configuration.ConfigurationSettings.AppSettings.Get("NotificationWeiXin") == "true")
                        {
                            String from = System.Configuration.ConfigurationSettings.AppSettings.Get("fromMail");
                            String to = System.Configuration.ConfigurationSettings.AppSettings.Get("toMail");
                            String host = System.Configuration.ConfigurationSettings.AppSettings.Get("host");
                            String user = System.Configuration.ConfigurationSettings.AppSettings.Get("username");
                            String pwd = System.Configuration.ConfigurationSettings.AppSettings.Get("password");
                            try
                            {
                                sendMail(host, user, pwd, from, to, message.ToString(), message.ToString());
                            }
                            catch (Exception ex)
                            {
                                Log_SMTP_Message("failed to send to weixin..");
                            }
                            
        
                        }

						// Insert "message" into DB
						Log_SMTP_Message("Received message:");
						Log_SMTP_Message(message.ToString());
					}
				}
			}

		}

		private System.IO.StreamWriter objLog = null;

		public void Log_Message(string logmessage)
		{
			if (objLog != null)
			{
				objLog.WriteLine(logmessage);
			}
		}

		private System.IO.StreamWriter objLogSMTP = null;

		public void Log_SMTP_Message(string logmessage)
		{
			if (objLogSMTP != null)
			{
				objLogSMTP.WriteLine(logmessage);
			}
		}

        static bool mailSent = false;
        public static void SendCompletedCallback(object sender, AsyncCompletedEventArgs e)
        {
            // Get the unique identifier for this asynchronous operation.
             String token = (string) e.UserState;
           
            if (e.Cancelled)
            {
                 Console.WriteLine("[{0}] Send canceled.", token);
            }
            if (e.Error != null)
            {
                 Console.WriteLine("[{0}] {1}", token, e.Error.ToString());
            } else
            {
                Console.WriteLine("Message sent.");
            }
            mailSent = true;
        }

        public static void sendMail(String host, String user, String pwd, String fromMail, String toMail, String subject, String strbody)
        {
            // Command line argument must the the SMTP host.
            SmtpClient client = new SmtpClient(host, 25);
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
            // Clean up.
            message.Dispose();
            Console.WriteLine("Goodbye.");
        }

	}
}
