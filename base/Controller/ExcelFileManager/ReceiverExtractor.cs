using System;
using System.Collections.Generic;
using Excel = Microsoft.Office.Interop.Excel;

using _base.Model;
using System.Linq;

namespace _base.Controller
{
	class ReceiverExtractor
	{
		private Dictionary<Receiver, Excel.Range> allReceiversInSheet = null;
		private Excel.Worksheet currentSheet = null;

		public List<Receiver> AllReceiversInSheet
		{
			get
			{
				if (allReceiversInSheet == null)
				{
					GetReceivers();
				}

				return allReceiversInSheet.Keys.ToList();
			}
		}
		// TODO вопрос между производительностью и памятью?
		//		к данному свойству будут обрачаться несколько раз,
		//		если данных мало, то не причин беспокоиться,
		//		если данных много, то каждый раз выполнять Linq-выражение снижает производительность программы,
		//		а если выбрать сохранение в памяти то требуется большой объем памяти
		//		что вы посоветуете? (я выбрал Linq-выражение вместо хранение в памяти)
		public List<string> NamesOfAllReceiversInSheet => AllReceiversInSheet.Select(u => u.Name).ToList();

		public ReceiverExtractor(Excel.Worksheet currentSheet)
		{
			this.currentSheet = currentSheet;
		}

		private void GetReceivers()
		{
			allReceiversInSheet = new Dictionary<Receiver, Excel.Range>();
			List<string> dataAfterStartTrigger;
			Receiver currentReceiver;

			foreach (Excel.Range oneDataRange in GetListOfDataRanges(currentSheet.UsedRange))
			{
				if ((dataAfterStartTrigger = GetValuesInRowAfterStartTriggerInDataRange(oneDataRange)) != null)
				{
					if ((currentReceiver = GetReceiverFromDataInRange(dataAfterStartTrigger)) != null)
					{
						if (IsReceiverNameUniq(currentReceiver) == false)
						{
							MakeReceiverNameUniq(currentReceiver);
						}

						allReceiversInSheet.Add(currentReceiver, oneDataRange);
					}
				}
			}
		}

		private List<Excel.Range> GetListOfDataRanges(Excel.Range TotalRangesInSheet)
		{
			List<Excel.Range> listOfDataRanges = null;

			try
			{
				FindCellByValue cellFinder = new FindCellByValue(TotalRangesInSheet);

				cellFinder.FindCells(GlobalObjects.Configuration.StartTrigger);
				List<Excel.Range> Start = cellFinder.FoundCells;

				cellFinder.FindCells(GlobalObjects.Configuration.EndTrigger);
				List<Excel.Range> End = cellFinder.FoundCells;

				ConverterFromStartAndEndCellsToRanges converter = new ConverterFromStartAndEndCellsToRanges();
				listOfDataRanges = converter.GetRanges(Start, End);
			}
			catch (ArgumentNullException)
			{ }

			return listOfDataRanges;
		}

		private List<string> GetValuesInRowAfterStartTriggerInDataRange(Excel.Range receiverDataRange)
		{
			int columnsCount = receiverDataRange.Columns.Count;
			List<string> listOfValues = new List<string>();

			Excel.Range currentCell = null;

			// currentCellsColumnIndex starts with 2 because it we skip the cell with trigger
			for (int currentCellsColumnIndex = 2; currentCellsColumnIndex <= columnsCount; currentCellsColumnIndex++)
			{
				currentCell = receiverDataRange.Item[1, currentCellsColumnIndex];
				if (currentCell.Text.Length > 0)
				{
					listOfValues.Add(currentCell.Text);
				}
			}

			return listOfValues.Count > 0 ? listOfValues : null;
		}

		private Receiver GetReceiverFromDataInRange(List<string> dataInRowAfterStartTrigger)
		{
			Receiver receiver = GetReceiverWithDataFromRowAfterStartTrigger(dataInRowAfterStartTrigger);

			if (DoesReceiverHaveAnyData(receiver) == false)
			{
				return null;
			}

			if (DoesReceiverHaveAllData(receiver) == false)
			{
				try
				{
					receiver = GetReceiverWhoseMissingDataIsFilled(receiver);
				}
				catch
				{
					return null;
				}
			}

			return receiver;
		}

		private bool DoesReceiverHaveAnyData(Receiver receiver)
		{
			return receiver.ID != 0 || string.IsNullOrEmpty(receiver.Name) == false || string.IsNullOrEmpty(receiver.EAddress) == false;
		}

		private bool DoesReceiverHaveAllData(Receiver receiver)
		{
			return receiver.ID != 0 && string.IsNullOrEmpty(receiver.Name) == false && string.IsNullOrEmpty(receiver.EAddress) == false;
		}

		private Receiver GetReceiverWithDataFromRowAfterStartTrigger(List<string> dataInRowAfterStartTrigger)
		{
			ReceiverManipulator receiverManipulator = new ReceiverManipulator();

			foreach (string item in dataInRowAfterStartTrigger)
			{
				if (ReceiverManipulator.TryGetId(item) != null)
				{
					receiverManipulator.Receiver.ID = (long)ReceiverManipulator.TryGetId(item);
				}
				else if (ReceiverManipulator.IsEAddress(item))
				{
					receiverManipulator.Receiver.EAddress = item;
				}
				else if (string.IsNullOrEmpty(receiverManipulator.Receiver.Name))
				{
					receiverManipulator.Receiver.Name = item;
				}
			}

			return receiverManipulator.Receiver;
		}

		private Receiver GetReceiverWhoseMissingDataIsFilled(Receiver receiver)
		{
			ReceiverManipulator receiverManipulator = new ReceiverManipulator();

			if (receiver.ID != 0)
			{
				if (receiverManipulator.Fill(receiver.ID) == false)
				{
					throw new Exception("Could not find specified receiver from DataBase to fill missing data");
				}
			}
			else if (string.IsNullOrEmpty(receiver.EAddress) == false)
			{
				if (!receiverManipulator.Fill(receiver.EAddress))
				{
					// if there is no data about receiver in dataBase and there is given only eAddress in Excel file,
					// accept it and send to this eAddress even if there is no enough data
					receiverManipulator.Fill(0, receiver.EAddress, receiver.EAddress);
				}
			}
			else if (string.IsNullOrEmpty(receiver.Name) == false)
			{
				if (receiverManipulator.Fill(receiver.Name) == false)
				{
					throw new Exception("Could not find specified receiver from DataBase to fill missing data");
				}
			}

			return receiverManipulator.Receiver;
		}

		private bool IsReceiverNameUniq(Receiver receiver)
		{
			return NamesOfAllReceiversInSheet.Contains(receiver.Name) == false;
		}

		private void MakeReceiverNameUniq(Receiver receiver)
		{
			int i = NamesOfAllReceiversInSheet.Count(u => u.StartsWith(receiver.Name));
			receiver.Name = receiver.Name + " (" + i + ")";
		}

		public bool IsAnyReceiverExtracted()
		{
			return AllReceiversInSheet != null && AllReceiversInSheet.Count > 0;
		}

		public Excel.Range GetReceiverDataRange(Receiver receiver)
		{
			if (allReceiversInSheet.Keys.Contains(receiver))
			{
				return allReceiversInSheet[receiver];
			}
			else
			{
				throw new ArgumentException($"Could not find receiver: {receiver.Name}");
			}
		}

		public Excel.Range GetReceiverDataRange(string receiverName)
		{
			Receiver receiver = allReceiversInSheet.Keys.FirstOrDefault(u => u.Name == receiverName);
			return GetReceiverDataRange(receiver);
		}
	}
}
