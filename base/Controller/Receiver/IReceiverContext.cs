using System.Collections.Generic;

using _base.Model;

namespace _base.Controller
{
	public interface IReceiverContext
	{
		bool Create(string Name, string EAddress);
		bool Update(long ID, string Name, string EAddress);
		bool Delete(long ID);
		List<Receiver> GetReceivers();

		public static IReceiverContext CreateReceiverContext()
		{
			return new ReceiverContext();
		}
	}
}
