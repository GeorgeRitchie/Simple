using System;

using MsOutlook = Microsoft.Office.Interop.Outlook;

using _base.Exceptions;

namespace _base.Controller
{
	class OutLookSendMail
	{
		MsOutlook.Application outlookApp = null;
		MsOutlook.MailItem mail = null;

		public bool SendMailWithOutlook(MailItem mailItem, MailSendType sendType)
		{
			try
			{
				OpenOutLookApp();
				CreateMailItem(mailItem);
				OpenMailSendingSessionForSpecifiedSender(mailItem);
				Send(sendType);
				Dispose();

				return true;
			}
			catch (Exception ex) when (!(ex is MailSendException))
			{
				Dispose();

				return false;
			}
		}

		private void OpenOutLookApp()
		{
			outlookApp = new MsOutlook.Application();
			if (outlookApp == null)
				throw new Exception("Could not open OutLook app.");
		}

		private void CreateMailItem(MailItem mailItem)
		{
			mail = (MsOutlook.MailItem)outlookApp.CreateItem(MsOutlook.OlItemType.olMailItem);

			mail.HTMLBody = mailItem.HtmlBody;

			if (mailItem.FilePaths != null)
			{
				foreach (string file in mailItem.FilePaths)
				{
					mail.Attachments.Add(file);
				}
			}

			mail.Subject = mailItem.Subject;
			mail.To = mailItem.Recipients;
		}

		private void OpenMailSendingSessionForSpecifiedSender(MailItem mailItem)
		{
			bool didFoundSpecifiedSender = false;

			MsOutlook.Accounts accounts = mail.Session.Accounts;
			for (int i = 1; i <= accounts.Count; i++)
			{
				string accountfoundEAddress = accounts[i].SmtpAddress.ToLower();
				if (mailItem.SenderEAddress.ToLower() == accountfoundEAddress)
				{
					mail.SendUsingAccount = accounts[i];
					MsOutlook.Recipient recipient = mail.Session.CreateRecipient(accountfoundEAddress);
					mail.Sender = recipient.AddressEntry;
					didFoundSpecifiedSender = true;
					break;
				}
			}

			if (!didFoundSpecifiedSender)
			{
				throw new MailSendException($"There is no account in OUTLOOK program with '{mailItem.SenderEAddress}' email address.\nSet correct email address that is logged in OUTLOOK application!\n");
			}
		}

		private void Send(MailSendType sendType)
		{
			if (sendType == MailSendType.SendDirect)
				mail.Send();
			else if (sendType == MailSendType.ShowModal)
				mail.Display(true);
			else if (sendType == MailSendType.ShowModeless)
				mail.Display(false);
		}

		private void Dispose()
		{
			mail = null;
			outlookApp = null;

			GC.Collect();
			GC.WaitForPendingFinalizers();
			GC.Collect();
			GC.WaitForPendingFinalizers();
		}
	}
}
