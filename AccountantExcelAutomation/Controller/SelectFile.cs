using Microsoft.Win32;

namespace AccountantExcelAutomation.Controller
{
	static class SelectFile
	{
		public static string Title { get; set; } = "Choose Excel file";
		public static string Filter { get; set; } = "Excel Files|*.xls;*.xls;*.xlt;*.xltm;*.xltx;*.xlsx;*.xlsm";


		#region form for selecting file

		public static string[] SelectFiles()
		{
			string[] filesPath = null;

			// create an instance of OpenFileDialog and set it
			OpenFileDialog openFile = new OpenFileDialog();
			openFile.Title = Title;
			openFile.Filter = Filter;
			openFile.CheckFileExists = true;
			openFile.CheckPathExists = true;

			// open OpenFileDialog
			openFile.ShowDialog();

			// save selected files path before deleting OpenFileDialog object
			filesPath = openFile.FileNames;

			return filesPath;
		}

		#endregion
	}
}
