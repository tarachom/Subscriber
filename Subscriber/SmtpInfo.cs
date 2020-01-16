using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Subscriber
{
	/// <summary>
	/// Клас для інформації по SMTP
	/// </summary>
	public class SmtpInfo
	{
		public SmtpInfo()
		{

		}

		public SmtpInfo(string smtpServer, string email, string pass, int port)
		{
			SmtpServer = smtpServer;
			Email = email;
			Pass = pass;
			Port = port;
		}

		public string SmtpServer { get; set; }

		public string Email { get; set; }

		public string Pass { get; set; }

		public int Port { get; set; }
	}
}
