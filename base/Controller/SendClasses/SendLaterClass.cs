using System;
using System.Collections.Generic;
using System.Linq;

using _base.Exceptions;
using _base.Model;

namespace _base.Controller
{
	class SendLaterClass
	{
		public void CreateDeferredSending(DataOfSendingMails_OfSelectedSheet data)
		{
			ValidateTimeOfDeferredSendingMails(data);
			CreateWindowsTask(data);
		}

		private void ValidateTimeOfDeferredSendingMails(DataOfSendingMails_OfSelectedSheet dataOfSendingMails)
		{
			if (dataOfSendingMails.MailsSendDateTime <= DateTime.Now)
			{
				throw new ArgumentException($"The time of sending the mail must be longer than the current time! Remember that this time is used to create task in Windows Task Scheduler.\n");
			}
		}

		private void CreateWindowsTask(DataOfSendingMails_OfSelectedSheet dataOfSendingMails)
		{
			List<WindowsTaskScheduler_Object> listOfTaskObjects = WindowsTaskScheduler_Manager.GetExistingTasks();

			WindowsTaskScheduler_Object taskObject = new WindowsTaskScheduler_Object()
			{
				DateTimeOfSendingMail = dataOfSendingMails.MailsSendDateTime,
				DataOfSendingMails_FileFullName = dataOfSendingMails.DataOfSendingMailsFileFullName
			};

			taskObject.TaskID = GetUniqIdToTaskObject(listOfTaskObjects);

			try
			{
				WindowsTaskScheduler_Manager.CreateTask(taskObject);
			}
			catch (Exception ex)
			{
				throw new WindowsTaskCreateException(ex.Message, taskObject.TaskID, taskObject.DateTimeOfSendingMail);
			}
			finally
			{
				listOfTaskObjects.Add(taskObject);
				WindowsTaskScheduler_Manager.UpdateFileOfWindowsTasks(listOfTaskObjects);
			}
		}

		private long GetUniqIdToTaskObject(List<WindowsTaskScheduler_Object> taskObjects)
		{
			if (taskObjects != null && taskObjects.Count > 0)
			{
				return taskObjects.Select(u => { return u.TaskID; }).Max() + 1;
			}
			else
			{
				return 1;
			}
		}
	}
}
