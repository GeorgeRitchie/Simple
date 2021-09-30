using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Core;
using _base.Model;

namespace _base.Controller
{
	class OpenExcelFile
	{
		public static Excel.Application Open(string fileName)
		{
			Excel.Application app = null;

			try
			{
				app = new Excel.Application();
				app.Visible = GlobalObjects.Properties.ExcelAppVisibility;
				app.AutomationSecurity = MsoAutomationSecurity.msoAutomationSecurityForceDisable;
				app.CSVDisplayNumberConversionWarning = false;
				app.Workbooks.Open(fileName, ReadOnly: true, Password: "");
			}
			catch (System.Runtime.InteropServices.COMException e)
			{
				Logger.Log($"File {fileName} processing failed! Exception was thrown. Exception message: {e.Message}", typeof(OpenExcelFile).FullName + '.' + nameof(Open));

				app?.Quit();
			}

			return app;
		}
	}
}
