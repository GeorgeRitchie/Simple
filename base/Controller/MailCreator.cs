using System;
using System.Collections.Generic;
using System.Text;
using _base.Model;
using System.Net.Mail;

namespace _base.Controller
{
	static class MailCreator
	{
		static public MailMessage Create(Mail_Object mail, string senderName, string senderEAddress)
		{
			// отправитель - устанавливаем адрес и отображаемое в письме имя
			MailAddress from = new MailAddress(senderEAddress, senderName);

			// кому отправляем
			MailAddress to = new MailAddress(mail.ReceiverEAddress);

			// создаем объект сообщения
			MailMessage m = new MailMessage(from, to);

			// тема письма
			m.Subject = mail.MailTitle;

			// текст письма
			m.Body = $"<h2>{mail.MailText}</h2>";

			// письмо представляет код html
			m.IsBodyHtml = true;

			// прикрепляем файл в письмо
			if (mail.AttachedFile != null && mail.AttachedFile.Length > 0)
			{
				m.Attachments.Add(new Attachment(mail.AttachedFile));
			}

			return m;
		}
	}
}
