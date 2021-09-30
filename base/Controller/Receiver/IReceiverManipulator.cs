using _base.Model;

namespace _base.Controller
{
	interface IReceiverManipulator
	{
		bool Fill(long ID);
		bool Fill(string Name);
		void Fill(long ID, string Name, string EAddress);

		Receiver Receiver { get; }
	}
}
