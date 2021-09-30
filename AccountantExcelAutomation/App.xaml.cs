using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Threading.Tasks;
using System.Windows;
using _base;
using _base.Controller;
using _base.Exceptions;
using _base.Model;
using AccountantExcelAutomation.View;

namespace AccountantExcelAutomation
{
	/// <summary>
	/// Interaction logic for App.xaml
	/// </summary>
	public partial class App : Application
	{
		App()
		{
			InitializeComponent();

		}

		private static void ClearAllGlobalResources()
		{
			if (GlobalObjects.Configuration != null)
			{
				GlobalObjects.Configuration = null;
			}
			if (GlobalObjects.Properties != null)
			{
				GlobalObjects.Properties = null;
			}

			GC.Collect();
			GC.WaitForPendingFinalizers();
			GC.Collect();
			GC.WaitForPendingFinalizers();
			GC.Collect();
			GC.WaitForPendingFinalizers();
		}


		private static bool ForceUserToSetUpProgramOptions(string[] args)
		{
			while ((GlobalObjects.Configuration.MailSenderEAddress == null || GlobalObjects.Configuration.MailSenderEAddress.Length == 0) && !args.Contains(GlobalObjects.Properties.ConsoleProgramKey_Configure))
			{
				SettingsWindow settingsWindow = new SettingsWindow(null);
				settingsWindow.ShowDialog();

				if (GlobalObjects.Configuration.MailSenderEAddress == null || GlobalObjects.Configuration.MailSenderEAddress.Length == 0)
				{
					if (MessageBox.Show("Program is not configured! Without configuration program will not work correctly.\nExit from program?", "ExcelAuto Program configuration error", MessageBoxButton.YesNo, MessageBoxImage.Warning, MessageBoxResult.No, MessageBoxOptions.ServiceNotification) == MessageBoxResult.Yes)
					{
						Logger.Log("Error!\nUser exited program without configuring it!\n", typeof(App) + "." + nameof(Main));
						return false;
					}
				}
			}

			return true;
		}

		private static int FindIDIndex()
		{
			int indexOfSendParam = 0;
			foreach (var item in args)
			{
				if (item == GlobalObjects.Properties.ConsoleProgramKey_Send)
				{
					break;
				}
				indexOfSendParam++;
			}

			return indexOfSendParam + 1;
		}

		private static void RunProgramWithSendParam(string[] args)
		{
			int indexOfIDParam = FindIDIndex();

			try
			{
				MainExcelController excelController = new MainExcelController();
				excelController.SendNow(Convert.ToInt64(args[indexOfIDParam]));
			}
			catch (MailSendException e)
			{
				Logger.Log("Mails are not sent\n" + e.Message, typeof(App) + "." + nameof(Main));
				MessageBox.Show("Mails are not sent\n" + e.Message, "ExcelAuto Error!", MessageBoxButton.OK, MessageBoxImage.Error, MessageBoxResult.None, MessageBoxOptions.ServiceNotification);
			}
			catch (Exception ex)
			{
				Logger.Log("Error occured while sending deffered mails\n" + ex.Message, typeof(App) + "." + nameof(Main));
			}
		}

		private static void RunProgramWithConfigureParam()
		{
			try
			{
				SettingsWindow settingsWindow = new SettingsWindow(null);
				app.Run(settingsWindow);
			}
			catch (Exception ex)
			{
				Logger.Log("Error occured while configuring the program\n" + ex.Message, typeof(App) + "." + nameof(RunProgramWithConfigureParam));
			}
		}

		private static string SelectFirstFileNameWithCorrectExtension()
		{
			foreach (var item in args)
			{
				FileInfo fileToCheck = new FileInfo(item);

				if (allowedExtensions.Contains(fileToCheck.Extension))
				{
					return item;
				}
			}

			return null;
		}


		private static void RunProgramWithGivenParams(string[] args)
		{
			if (args.Contains(GlobalObjects.Properties.ConsoleProgramKey_Send))
			{
				RunProgramWithSendParam(args);
			}
			else if (args.Contains(GlobalObjects.Properties.ConsoleProgramKey_Configure))
			{
				RunProgramWithConfigureParam();
			}
			else
			{
				string excelFile = SelectFirstFileNameWithCorrectExtension();

				if (excelFile != null && excelFile.Length > 0)
				{
					RunMainWindow(new MainWindow(excelFile));
				}
			}
		}

		private static void RunMainWindow(MainWindow window)
		{
			try
			{
				if (window.CanRunProgram)
				{
					app.Run(window);
				}
			}
			catch (Exception ex)
			{
				Logger.Log(ex.Message, "Application.Run");
			}
		}

		private static string[] allowedExtensions = { ".xls", ".xlt", ".xltm", ".xltx", ".xlsx", ".xlsm" };
		private static MainWindow window = null;
		private static App app = new App();
		private static string[] args = null;

		[STAThread]
		static void Main(string[] args)
		{

			Logger.Log("APPLICATION LAUNCHED", "");

			if (ForceUserToSetUpProgramOptions(args) == false)
			{
				Logger.Log("User did not set up program settings", "");
				Logger.Log("SESSION ENDED", "");
				ClearAllGlobalResources();
				return;
			}

			if (args.Length > 0)
			{
				App.args = args;
				RunProgramWithGivenParams(args);
			}
			else
			{
				RunMainWindow(new MainWindow());
			}

			Logger.Log("SESSION ENDED", "");

			window = null;

			ClearAllGlobalResources();
		}
	}
}
