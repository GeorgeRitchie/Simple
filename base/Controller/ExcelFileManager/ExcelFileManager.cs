using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;

using Excel = Microsoft.Office.Interop.Excel;
using _base.Model;

namespace _base.Controller
{
	class ExcelFileManager : IExcelFileManager, IDisposable
	{
		private Excel.Application app;
		private Excel.Worksheet worksheet;
		private ReceiverExtractor receiverExtractorForSelectedSheet = null;
		private DataOfSendingMails_OfSelectedSheet data;

		public ExcelFileManager(Excel.Application appWithOpenedWorkbook)
		{
			app = appWithOpenedWorkbook;
		}

		// #1
		public List<string> GetWorkbookSheetsNames()
		{
			List<string> totalSheetsName = GetSheetsNames();
			DeleteSheetsNamesThatDoNotHaveTriggers(totalSheetsName);
			List<string> validSheetsNames = GetSheetsNamesThatHaveTriggersWithReceiversData(totalSheetsName);

			return validSheetsNames.Count > 0 ? validSheetsNames : null;
		}

		private List<string> GetSheetsNames()
		{
			List<string> sheetsName = new List<string>();

			foreach (Excel.Worksheet worksheet in app.Workbooks[1].Worksheets)
			{
				sheetsName.Add(worksheet.Name);
			}

			return sheetsName;
		}

		private void DeleteSheetsNamesThatDoNotHaveTriggers(List<string> totalSheetsNames)
		{
			foreach (Excel.Worksheet currentSheet in app.Worksheets)
			{
				Excel.Range totalRangesOfCurrentSheet = currentSheet.UsedRange;
				FindCellByValue cellFinder = new FindCellByValue(totalRangesOfCurrentSheet);
				cellFinder.FindCells(GlobalObjects.Configuration.StartTrigger);

				if (cellFinder.IsAnyCellFound() == false)
				{
					totalSheetsNames.Remove(currentSheet.Name);
				}
			}
		}

		private List<string> GetSheetsNamesThatHaveTriggersWithReceiversData(List<string> totalSheetsNames)
		{
			List<string> sheetsNameWithTriggersAndReceiversData = new List<string>();

			foreach (string item in totalSheetsNames)
			{
				ReceiverExtractor receiverExtractor = new ReceiverExtractor(app.Workbooks[1].Worksheets[item]);
				
				if (receiverExtractor.IsAnyReceiverExtracted())
				{
					sheetsNameWithTriggersAndReceiversData.Add(item);
				}
			}

			return sheetsNameWithTriggersAndReceiversData;
		}

		// #2
		public void SetChosenSheetName(string SheetName)
		{
			if (string.IsNullOrEmpty(SheetName))
			{
				throw new ArgumentException($"Invalid passed argument {nameof(SheetName)}");
			}

			worksheet = app.Workbooks[1].Worksheets[SheetName];
			receiverExtractorForSelectedSheet = new ReceiverExtractor(worksheet);
		}

		// #3
		public List<string> GetReceiversNames()
		{
			return receiverExtractorForSelectedSheet.NamesOfAllReceiversInSheet;
		}

		// #4
		public void SetSelectedReceivers(List<string> chosenReceivers)
		{
			FillDataOfSendingMails(chosenReceivers);
		}

		private void FillDataOfSendingMails(List<string> chosenReceivers)
		{
			data = new DataOfSendingMails_OfSelectedSheet();
			data.ExcelFileName = app.Workbooks[1].Name;
			data.SelectedSheetName = worksheet.Name;
			data.Mails = CreateMailObjectsForEachChosenReceiver(chosenReceivers);
		}

		private List<Mail_Object> CreateMailObjectsForEachChosenReceiver(List<string> chosenReceivers)
		{
			List<Mail_Object> mailObjects = null;

			if (chosenReceivers.Count > 0)
			{
				mailObjects = new List<Mail_Object>();

				foreach (string currentReceiverName in chosenReceivers)
				{
					Receiver currentReceiver = receiverExtractorForSelectedSheet.AllReceiversInSheet.Find(u => u.Name == currentReceiverName);

					Mail_Object mail = new Mail_Object()
					{
						ReceiverID = currentReceiver.ID,
						ReceiverName = currentReceiverName,
						ReceiverEAddress = currentReceiver.EAddress,
						DataRangeOfReceiver = receiverExtractorForSelectedSheet.GetReceiverDataRange(currentReceiver),
						MailTitle = GlobalObjects.Configuration.MailTitle,
						MailText = GlobalObjects.Configuration.MailText
					};

					mailObjects.Add(mail);
				}
			}

			return mailObjects;
		}

		// #5
		public DataOfSendingMails_OfSelectedSheet MakeMailObjectsFromReceiversData()
		{
			data.MailDataRecorder.CreateDirectory();
			CreateAttechedFilesForEachReceiver();
			data.MailDataRecorder.RecordAllData();

			return data;
		}

		private void CreateAttechedFilesForEachReceiver()
		{
			AttachedFilesCreator filesCreator = new AttachedFilesCreator();
			foreach (Mail_Object item in data.Mails)
			{
				filesCreator.CreateAttachedFile(item.DataRangeOfReceiver, item.AttachedFile_FullName, GlobalObjects.Configuration.ChosenAttachedFileExtentionAsInt);
			}
		}

		#region Dispose Interface implementation

		private bool disposed = false;

		public void Dispose()
		{
			Dispose(true);
			GC.SuppressFinalize(this);
		}

		protected virtual void Dispose(bool disposing)
		{
			if (!disposed)
			{
				if (disposing)
				{
					// Освобождаем управляемые ресурсы
				}

				GC.Collect();
				GC.WaitForPendingFinalizers();

				// prevent closing app if it already closed or even not opened
				if (app != null && app.Workbooks.Count > 0)
				{
					app.Workbooks[1].Close(SaveChanges: false);
					app.Quit();
				}
				// if app is opened but there is no workbook opened
				else if (app != null)
				{
					app.Quit();
				}

				if (worksheet != null)
					Marshal.FinalReleaseComObject(worksheet);
				if (app != null)
					Marshal.FinalReleaseComObject(app);


				GC.Collect();
				GC.WaitForPendingFinalizers();

				disposed = true;
			}
		}

		~ExcelFileManager()
		{
			Dispose(false);
		}

		#endregion
	}
}
