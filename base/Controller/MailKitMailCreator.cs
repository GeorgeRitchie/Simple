using System;
using System.Collections.Generic;
using System.Text;
using MimeKit;
using MailKit.Net.Smtp;
using _base.Model;

namespace _base.Controller
{
	class MailKitMailCreator
	{
		static public MimeMessage Create(Mail_Object mail, string senderName, string senderEAddress)
		{
			MimeMessage emailMessage = new MimeMessage();
			
			emailMessage.From.Add(new MailboxAddress(senderName, senderEAddress));
			emailMessage.To.Add(new MailboxAddress(mail.ReceiverName, mail.ReceiverEAddress));
			emailMessage.Subject = mail.MailTitle;

			BodyBuilder emailBody = new BodyBuilder();
			emailBody.HtmlBody = $"<h2>{mail.MailText}</h2>";

			if (mail.AttachedFile != null && mail.AttachedFile.Length > 0)
			{
				emailBody.Attachments.Add(mail.AttachedFile);
			}

			emailMessage.Body = emailBody.ToMessageBody();

			return emailMessage;
		}
	}
}
