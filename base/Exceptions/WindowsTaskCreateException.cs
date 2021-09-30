using System;

namespace _base.Exceptions
{
	public class WindowsTaskCreateException : Exception
	{
		public long TaskID { get; private set; }
		public DateTime DateTimeOfSendingMail { get; private set; }

		public WindowsTaskCreateException(string msg, long TaskID, DateTime DateTimeOfSendingMailOfTask) : base(msg)
		{
			this.TaskID = TaskID;
			this.DateTimeOfSendingMail = DateTimeOfSendingMailOfTask;
		}
	}
}
