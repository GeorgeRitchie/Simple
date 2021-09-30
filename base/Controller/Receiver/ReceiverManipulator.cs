using System.Collections.Generic;
using _base.Model;
using System.Text.RegularExpressions;

namespace _base.Controller
{
	public class ReceiverManipulator : IReceiverManipulator
	{
		readonly static string pattern = @"^(?("")(""[^""]+?""@)|(([0-9a-z]((\.(?!\.))|[-!#\$%&'\*\+/=\?\^`\{\}\|~\w])*)(?<=[0-9a-z])@))" +
				@"(?(\[)(\[(\d{1,3}\.){3}\d{1,3}\])|(([0-9a-z][-\w]*[0-9a-z]*\.)+[a-z0-9]{2,17}))$";

		private Receiver receiver = new Receiver();
		private IReceiverContext receiverContext = null;

		public ReceiverManipulator()
		{
			receiverContext = new ReceiverContext();
		}

		public Receiver Receiver { get => receiver; }

		public bool Fill(long ID)
		{
			try
			{
				List<Receiver> receivers = receiverContext.GetReceivers();
				
				return Fill(receivers.Find(u => u.ID == ID));
			}
			catch
			{
				return false;
			}
		}

		public bool Fill(string Name)
		{
			try
			{
				List<Receiver> receivers = receiverContext.GetReceivers();

				return Fill(receivers.Find(u => u.Name == Name));
			}
			catch
			{
				return false;
			}
		}

		public void Fill(long ID, string Name, string EAddress)
		{
			receiver.ID = ID;
			receiver.Name = Name;
			receiver.EAddress = EAddress;
		}

		private bool Fill(Receiver receiver)
		{
			if (receiver != null)
			{
				receiver.Name = receiver.Name;
				receiver.EAddress = receiver.EAddress;
				receiver.ID = receiver.ID;

				return true;
			}

			return false;
		}

		public static bool IsEAddress(string str)
		{
			return Regex.IsMatch(str, pattern, RegexOptions.IgnoreCase);
		}

		public static long? TryGetId(string str)
		{
			long id = 0;
			return long.TryParse(str, out id) ? (long?)id : null;
		}
	}
}
