using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.Json;

using _base.Model;

namespace _base.Controller
{
	class ReceiverContext : IReceiverContext
	{
		private readonly string ReceiversFilePath = GlobalObjects.Properties.ConfigurationFilesDirectoryName + "\\Receivers.json";

		bool IReceiverContext.Create(string Name, string EAddress)
		{
			try
			{
				List<Receiver> receivers = GetReceiversFromFile();
				long id = GetUniqId(receivers);
				Receiver receiver = new Receiver() { EAddress = EAddress, Name = Name, ID = id };
				receivers.Add(receiver);
				WriteReceiversToFile(receivers);
			}
			catch
			{
				return false;
			}

			return true;
		}

		private List<Receiver> GetReceiversFromFile()
		{
			List<Receiver> receivers = null;

			using (StreamReader reader = new StreamReader(new FileStream(ReceiversFilePath, FileMode.OpenOrCreate)))
			{
				string tempStr = reader.ReadToEnd();
				if (!string.IsNullOrEmpty(tempStr))
					receivers = JsonSerializer.Deserialize<List<Receiver>>(tempStr, new JsonSerializerOptions() { WriteIndented = true });
			}

			if (receivers == null)
			{
				receivers = new List<Receiver>();
			}

			return receivers;
		}

		private long GetUniqId(List<Receiver> receivers)
		{
			long id = 1;

			if (receivers != null && receivers.Count > 0)
			{
				long lastId = receivers.Select(u => { return u.ID; }).Max();

				if (IsAnyIdMissed(lastId, receivers) == false)
				{
					id = GetNextId(lastId);
				}
				else
				{
					id = GetFirstMissedId(lastId, receivers);
				}
			}

			return id;
		}

		private bool IsAnyIdMissed(long lastId, List<Receiver> receivers)
		{
			// if any id is not missed count and last id must be same value
			return lastId == receivers.Count;
		}

		private long GetNextId(long lastId)
		{
			return lastId + 1;
		}

		private long GetFirstMissedId(long lastId, List<Receiver> receivers)
		{
			long[] eachNumbersToLastID = new long[lastId];

			for (int i = 0; i < eachNumbersToLastID.Length; i++)
			{
				eachNumbersToLastID[i] = i + 1;
			}

			List<long> allNumbersInReceiversID = receivers.Select(u => { return u.ID; }).ToList();

			List<long> skipedIds = eachNumbersToLastID.Except(allNumbersInReceiversID).ToList();

			return skipedIds.FirstOrDefault();
		}

		private void WriteReceiversToFile(List<Receiver> receivers)
		{
			if (receivers == null)
			{
				throw new ArgumentNullException();
			}

			string tempStr = JsonSerializer.Serialize(receivers, new JsonSerializerOptions() { WriteIndented = true });

			using (StreamWriter writer = new StreamWriter(new FileStream(ReceiversFilePath, FileMode.Create)))
			{
				writer.Write(tempStr);
			}
		}

		bool IReceiverContext.Update(long ID, string Name, string EAddress)
		{
			try
			{
				List<Receiver> receivers = GetReceiversFromFile();
				Receiver tempReceiver = receivers.Find(u => u.ID == ID);
				tempReceiver.Name = Name;
				tempReceiver.EAddress = EAddress;
				WriteReceiversToFile(receivers);
			}
			catch
			{
				return false;
			}

			return true;
		}

		bool IReceiverContext.Delete(long ID)
		{
			try
			{
				List<Receiver> receivers = GetReceiversFromFile();
				receivers.Remove(receivers.Find(u => u.ID == ID));
				WriteReceiversToFile(receivers);
			}
			catch
			{
				return false;
			}

			return true;
		}

		List<Receiver> IReceiverContext.GetReceivers()
		{
			return GetReceiversFromFile();
		}
	}
}
