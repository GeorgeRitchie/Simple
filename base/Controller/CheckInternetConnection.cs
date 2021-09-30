using System.Net.NetworkInformation;

namespace _base.Controller
{
	class CheckInternetConnection
	{
		public static bool IsConnectedToInternet()
		{
			IPStatus[] status = ConnectToInternetAndReturnStatus();

			foreach (IPStatus item in status)
			{
				if (item == IPStatus.Success)
				{
					return true;
				}
			}

			return false;
		}

		private static IPStatus[] ConnectToInternetAndReturnStatus()
		{
			IPStatus[] status = new IPStatus[3] { IPStatus.Unknown, IPStatus.Unknown, IPStatus.Unknown };

			try
			{
				status[0] = TryToConnectToSite(@"google.com");
				status[1] = TryToConnectToSite(@"microsoft.com");
				status[2] = TryToConnectToSite(@"amazon.com");
			}
			catch { }

			return status;
		}

		private static IPStatus TryToConnectToSite(string siteAddress)
		{
			Ping p = new Ping();
			PingReply pingreply = p.Send(siteAddress);
			return pingreply.Status;
		}
	}
}
