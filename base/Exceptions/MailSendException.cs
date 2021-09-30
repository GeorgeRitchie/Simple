using System;

namespace _base.Exceptions
{
	public class MailSendException :Exception
	{
		// this Exception must be catch only by main method (method who starts the program or GUI)
		public MailSendException(string msg) :base (msg)
		{

		}
	}
}
