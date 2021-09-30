using System;
using System.Collections.Generic;
using System.Text.Json;
using System.IO;
using System.Linq;

using _base.Model;

namespace _base.Controller
{
	class SendMail
	{
		public void SendNow(DataOfSendingMails_OfSelectedSheet data)
		{
			ValidateDataOfSendingMails(data);

			SendNowClass sendNow = new SendNowClass();
			sendNow.SendNow(data);

			UpdateFileOfDataOfSendingMails(data);
		}

		private void ValidateDataOfSendingMails(DataOfSendingMails_OfSelectedSheet dataOfSendingMails)
		{
			if (dataOfSendingMails == null || dataOfSendingMails.Mails == null || dataOfSendingMails.Mails.Count == 0)
			{
				throw new ArgumentException($"Argument {nameof(dataOfSendingMails)} can not be null and must have any object in {nameof(dataOfSendingMails.Mails)} propertry\n");
			}

			if (dataOfSendingMails.AreMailsSent)
			{
				throw new InvalidOperationException($"Mails in given parameter {nameof(dataOfSendingMails)} have already sent!");
			}
		}

		private void UpdateFileOfDataOfSendingMails(DataOfSendingMails_OfSelectedSheet dataOfSendingMails)
		{
				dataOfSendingMails.MailDataRecorder.RecordAllData();
		}

		public void CreateDeferredSending(DataOfSendingMails_OfSelectedSheet data)
		{

			try
			{
				ValidateDataOfSendingMails(data);

				SendLaterClass sendLater = new SendLaterClass();
				// this code can throw an exception that must not be catched by this method
				sendLater.CreateDeferredSending(data);
			}
			finally
			{
				UpdateFileOfDataOfSendingMails(data);
			}
		}

		public void SendDeferredSending(long id)
		{
			List<WindowsTaskScheduler_Object> listOfTaskObjects = WindowsTaskScheduler_Manager.GetExistingTasks();

			WindowsTaskScheduler_Object specifiedTask = listOfTaskObjects.Where(u => u.TaskID == id).FirstOrDefault();

			if (specifiedTask == null)
			{
				throw new InvalidOperationException($"There is no Task such: {id}\n");
			}

			DataOfSendingMails_OfSelectedSheet dataOfSendingMailsOfSpecifiedTask = GetDataOfSendingMailsOfSpecifiedTask(specifiedTask);

			ThrowExceptionIfTaskIsTryingToSendOthersMails(specifiedTask, dataOfSendingMailsOfSpecifiedTask);

			SendNow(dataOfSendingMailsOfSpecifiedTask);

			listOfTaskObjects.Remove(specifiedTask);
			WindowsTaskScheduler_Manager.UpdateFileOfWindowsTasks(listOfTaskObjects);
		}

		private DataOfSendingMails_OfSelectedSheet GetDataOfSendingMailsOfSpecifiedTask(WindowsTaskScheduler_Object specifiedTask)
		{
			DataOfSendingMails_OfSelectedSheet dataOfSendingMails = null;

			using (StreamReader strRead = new StreamReader(specifiedTask.DataOfSendingMails_FileFullName))
			{
				string str = strRead.ReadToEnd();

				dataOfSendingMails = JsonSerializer.Deserialize<DataOfSendingMails_OfSelectedSheet>(str, new JsonSerializerOptions() { WriteIndented = true });
			}

			return dataOfSendingMails;
		}

		private void ThrowExceptionIfTaskIsTryingToSendOthersMails(WindowsTaskScheduler_Object specifiedTask, DataOfSendingMails_OfSelectedSheet dataOfSendingMails)
		{
			if (specifiedTask.DateTimeOfSendingMail != dataOfSendingMails.MailsSendDateTime)
			{
				string errorMsg = $"There is discrepancy between the date of sending '{dataOfSendingMails.MailsSendDateTime}' and the date of task {specifiedTask.DateTimeOfSendingMail}, it is not allowed ReMailing mails!\n";
				Logger.Log(errorMsg, typeof(SendMail).FullName + "." + nameof(ThrowExceptionIfTaskIsTryingToSendOthersMails));
				throw new InvalidOperationException(errorMsg);
			}
		}
	}
}
