using _base.Model;
using System.Collections.Generic;

namespace _base.Controller
{
	interface IExcelFileManager
	{
		public List<string> GetWorkbookSheetsNames();
		public void SetChosenSheetName(string SheetName);
		public List<string> GetReceiversNames();
		public void SetSelectedReceivers(List<string> chosenReceivers);
		public DataOfSendingMails_OfSelectedSheet MakeMailObjectsFromReceiversData();
	}
}
