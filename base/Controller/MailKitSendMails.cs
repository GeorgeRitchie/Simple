using System;
using System.Collections.Generic;
using System.Text;
using MimeKit;
using MailKit.Net.Smtp;
using _base.Model;

namespace _base.Controller
{
	class MailKitSendMails
	{
		public static void SendMail(Mail_Object mail)
		{
			using (var client = new SmtpClient())
			{
				client.Connect(GlobalObjects.Configuration.SMTP, GlobalObjects.Configuration.Port, false);
				client.Authenticate(GlobalObjects.Configuration.EAddress, GlobalObjects.Configuration.Password);
				client.Send(MailKitMailCreator.Create(mail, GlobalObjects.Configuration.Name, GlobalObjects.Configuration.EAddress));
				client.Disconnect(true);
			}
		}
	}
}
