using System;

using _base.Exceptions;
using _base.Model;

namespace _base.Controller
{
	class SendNowClass
	{
		public void SendNow(DataOfSendingMails_OfSelectedSheet data)
		{
			ValidateInternetConnection();
			SendMails(data);
			UpdateSendingMailsStatusToSent(data);
			NotifySenderAboutSuccessfulSending(data);
		}

		private void ValidateInternetConnection()
		{
			if (!CheckInternetConnection.IsConnectedToInternet())
			{
				Logger.Log("There is no internet connection to send mails.\n", typeof(SendMail) + "." + nameof(ValidateInternetConnection));
				throw new MailSendException("There is no internet connection to send mails.\n");
			}
		}

		private void SendMails(DataOfSendingMails_OfSelectedSheet dataOfSendingMails)
		{
			try
			{
				foreach (var receiver in dataOfSendingMails.Mails)
				{
					MailItem mailItem = new MailItem()
					{
						SenderEAddress = GlobalObjects.Configuration.MailSenderEAddress,
						Subject = receiver.MailTitle,
						HtmlBody = receiver.MailText,
						Recipients = receiver.ReceiverEAddress,
						FilePaths = new string[] { receiver.AttachedFile_FullName }
					};
					OutLookSendMail senderManager = new OutLookSendMail();
					senderManager.SendMailWithOutlook(mailItem, MailSendType.SendDirect);
				}
			}
			catch (Exception e)
			{
				Logger.Log("Some error occurred while sending mails:\n" + e.Message, typeof(SendMail) + "." + nameof(SendMails));
				throw new MailSendException("Some error occurred while sending mails.\n" + e.Message);
			}
		}

		private void UpdateSendingMailsStatusToSent(DataOfSendingMails_OfSelectedSheet dataOfSendingMails)
		{
			dataOfSendingMails.MailsSendDateTime = DateTime.Now;
			dataOfSendingMails.AreMailsSent = true;
		}

		private void NotifySenderAboutSuccessfulSending(DataOfSendingMails_OfSelectedSheet dataOfSendingMails)
		{
			string receiversNameAndEaddress = GetNameAndEAddressOfReceivers(dataOfSendingMails);

			try
			{
				MailItem mailItem = new MailItem()
				{
					SenderEAddress = GlobalObjects.Configuration.MailSenderEAddress,
					Subject = "Mails are sent",
					HtmlBody = $"Mails from sheet {dataOfSendingMails.SelectedSheetName} from file {dataOfSendingMails.ExcelFileName} are sent to \n {receiversNameAndEaddress}",
					Recipients = GlobalObjects.Configuration.MailSenderEAddress,
					FilePaths = null
				};
				OutLookSendMail senderManager = new OutLookSendMail();
				senderManager.SendMailWithOutlook(mailItem, MailSendType.SendDirect);
			}
			catch (Exception e)
			{
				Logger.Log("Some error occurred while sending mails:\n" + e.Message, typeof(SendMail) + "." + nameof(NotifySenderAboutSuccessfulSending));
				throw new MailSendException("Some error occurred while sending mails.\n" + e.Message);
			}
		}

		private string GetNameAndEAddressOfReceivers(DataOfSendingMails_OfSelectedSheet dataOfSendingMails)
		{
			string receiversNameAndEaddress = "<br/>";
			int index = 1;
			foreach (var item in dataOfSendingMails.Mails)
			{
				// write each receiver's name and eaddress in new line in format "index	name eaddress"
				receiversNameAndEaddress += index++ + "\t" + item.ReceiverName + "\t" + item.ReceiverEAddress + "\n<br/>";
			}

			return receiversNameAndEaddress;
		}
	}
}
