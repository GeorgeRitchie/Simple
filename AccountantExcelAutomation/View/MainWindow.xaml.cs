using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;

using _base;
using _base.Controller;
using _base.Exceptions;
using _base.Model;

using AccountantExcelAutomation.Model;

namespace AccountantExcelAutomation.View
{
	/// <summary>
	/// Логика взаимодействия для MainWindow.xaml
	/// </summary>
	public partial class MainWindow : Window
	{
		private MainExcelController excelController;
		private string fileToOpen = null;
		public MainWindow(string file = null)
		{
			fileToOpen = file;

			excelController = new MainExcelController();

			InitializeComponent();
		}

		public bool CanRunProgram = true;
		string choosenFile = null;

		private void Window_Initialized(object sender, EventArgs e)
		{

			if (fileToOpen == null)
			{
				string[] files = Controller.SelectFile.SelectFiles();
				if (files != null && files.Length > 0)
				{
					choosenFile = files[0];
				}
			}
			else
			{
				choosenFile = fileToOpen;
			}


			if (choosenFile != null && choosenFile.Length > 0)
			{
				Sheets.ItemsSource = excelController.GetWorkbookSheetsNames(choosenFile);
			}
			else
			{
				Logger.Log("No file selected", "MainWindow");
				CanRunProgram = false;
			}
		}

		private void Sheets_SelectionChanged(object sender, SelectionChangedEventArgs e)
		{
			if (Sheets.SelectedIndex >= 0)
			{
				excelController.SetChoosenSheetName(Sheets.SelectedItem as string);

				List<string> Receivers = excelController.GetReceiversNames();

				if (Receivers.Contains("All"))
				{
					Receivers[Receivers.FindIndex(u => u == "All")] = "_All_";
				}

				Receivers.Insert(0, "All");

				if (Receivers.Count > 1)
				{
					this.Receivers = new List<ReceiverListBoxItem>();
					foreach (var item in Receivers)
					{
						this.Receivers.Add(new ReceiverListBoxItem() { Name = item, IsChecked = true });
						ReceiverListBoxItem.CheckedCount++;
					}
					ReceiversList.ItemsSource = this.Receivers;
					CheckboxFirstPreviousState = this.Receivers[0].IsChecked;

					SendLater.IsEnabled = true;
					SendNow.IsEnabled = true;
				}
			}
			else
			{
				ReceiversList.ItemsSource = this.Receivers = null;
			}
		}

		private List<ReceiverListBoxItem> Receivers;
		private bool? CheckboxFirstPreviousState;
		private void CheckBox_Click(object sender, RoutedEventArgs e)
		{
			CheckBox checkBox = sender as CheckBox;

			if (checkBox.Content.ToString() != "All")
			{
				if (checkBox.IsChecked == false)
				{
					ReceiverListBoxItem.CheckedCount--;
					if (ReceiverListBoxItem.CheckedCount > 1)
					{
						CheckboxFirstPreviousState = Receivers[0].IsChecked = null;
					}
					else
					{
						CheckboxFirstPreviousState = Receivers[0].IsChecked = false;
					}
				}
				else if (checkBox.IsChecked == true)
				{
					ReceiverListBoxItem.CheckedCount++;

					if (ReceiverListBoxItem.CheckedCount > 1 && ReceiverListBoxItem.CheckedCount < Receivers.Count)
					{
						CheckboxFirstPreviousState = Receivers[0].IsChecked = null;
					}
					else if (ReceiverListBoxItem.CheckedCount == Receivers.Count)
					{
						CheckboxFirstPreviousState = Receivers[0].IsChecked = true;
					}
				}
			}
			else if (checkBox.Content.ToString() == "All")
			{
				if (CheckboxFirstPreviousState == true)
				{
					CheckboxFirstPreviousState = checkBox.IsChecked = false;
					foreach (var item in Receivers)
					{
						if (item.Name != "All")
							item.IsChecked = false;
					}
					ReceiverListBoxItem.CheckedCount = 1;
				}
				else if (CheckboxFirstPreviousState == null || CheckboxFirstPreviousState == false)
				{
					CheckboxFirstPreviousState = checkBox.IsChecked = true;
					foreach (var item in Receivers)
					{
						if (item.Name != "All")
							item.IsChecked = true;
					}
					ReceiverListBoxItem.CheckedCount = Receivers.Count;
				}
			}
		}

		private List<string> GetSelectedReceivers(List<ReceiverListBoxItem> receiverLists)
		{
			List<string> selectedReceivers = null;

			if (ReceiverListBoxItem.CheckedCount > 1)
			{
				selectedReceivers = new List<string>();

				foreach (var item in receiverLists)
				{
					if (item.IsChecked == true && item.Name != "All")
					{
						if (item.Name == "_All_")
							selectedReceivers.Add("All");
						else
							selectedReceivers.Add(item.Name);
					}
				}
			}

			return selectedReceivers;
		}

		private void SendLater_Click(object sender, RoutedEventArgs e)
		{
			List<string> selectedReceivers = GetSelectedReceivers(Receivers);
			if (selectedReceivers == null || selectedReceivers.Count == 0)
			{
				MessageBox.Show("At least one receiver must be checked!", "Excel Auto Warning!", MessageBoxButton.OK, MessageBoxImage.Warning, MessageBoxResult.None, MessageBoxOptions.ServiceNotification);
				return;
			}

			SelectedReceivers = selectedReceivers;

			MainOperationsGrid.Visibility = Visibility.Hidden;
			SendLaterGrid.Visibility = Visibility.Visible;
		}

		private void SendNow_Click(object sender, RoutedEventArgs e)
		{
			List<string> selectedReceivers = GetSelectedReceivers(Receivers);
			if (selectedReceivers == null || selectedReceivers.Count == 0)
			{
				MessageBox.Show("At least one receiver must be checked!", "Excel Auto Warning!", MessageBoxButton.OK, MessageBoxImage.Warning, MessageBoxResult.None, MessageBoxOptions.ServiceNotification);
				return;
			}

			MainGrid.Cursor = Cursors.Wait;

			excelController.SetSelectedReceivers(selectedReceivers);
			excelController.MakeMailObjectsFromReceiversData();

			try
			{
				excelController.SendNow();
			}
			catch (MailSendException ex)
			{
				Logger.Log("Mails are not sent\n" + ex.Message, typeof(MainWindow) + "." + nameof(SendNow));
				MessageBox.Show("Mails are not sent\n" + ex.Message, "Excel Auto Error!", MessageBoxButton.OK, MessageBoxImage.Error, MessageBoxResult.None, MessageBoxOptions.ServiceNotification);
			}

			MainGrid.Cursor = null;

			this.Close();
		}

		private void Settings_Click(object sender, RoutedEventArgs e)
		{
			SettingsWindow settingsWindow = new SettingsWindow(this);
			settingsWindow.ShowDialog();

			reloadTheOpenedFile.Visibility = Visibility.Visible;
			Grid.SetColumnSpan(Sheets, 1);
		}

		private void Cancel_Click(object sender, RoutedEventArgs e)
		{
			MainOperationsGrid.Visibility = Visibility.Visible;
			SendLaterGrid.Visibility = Visibility.Hidden;
		}

		List<string> SelectedReceivers = null;
		private void Confirm_Click(object sender, RoutedEventArgs e)
		{
			DateTime? sendTime = DateOfSending.Value;
			if (sendTime == null || sendTime <= DateTime.Now)
			{
				MessageBox.Show("Choose Date and Time after current!");
				return;
			}

			MainGrid.Cursor = Cursors.Wait;

			try
			{
				excelController.SetSelectedReceivers(SelectedReceivers);
				excelController.MakeMailObjectsFromReceiversData();

				try
				{
					excelController.SendLater((DateTime)sendTime);
				}
				catch (WindowsTaskCreateException ex)
				{
					Logger.Log("Error occured during creating Windows Task" + ex.Message, typeof(MainWindow) + "." + nameof(Confirm_Click));

					WindowsTaskCreateExceptionNotification notification = new WindowsTaskCreateExceptionNotification(this, Process.GetCurrentProcess().MainModule.FileName, GlobalObjects.Properties.ConsoleProgramKey_Send + " " + ex.TaskID, ex.DateTimeOfSendingMail);
					notification.ShowDialog();
				}
			}
			catch (MailSendException ex)
			{
				Logger.Log("Mails are not sent\n" + ex.Message, typeof(MainWindow) + "." + nameof(SendNow));
				MessageBox.Show("Mails are not sent\n" + ex.Message, "Excel Auto Error!", MessageBoxButton.OK, MessageBoxImage.Error, MessageBoxResult.None, MessageBoxOptions.ServiceNotification);
			}

			MainGrid.Cursor = null;

			this.Close();
		}

		private void Window_Loaded(object sender, RoutedEventArgs e)
		{
			this.Focus();
			this.Activate();
		}

		private void Window_Closed(object sender, EventArgs e)
		{
			excelController?.CloseApp();
		}

		private void reloadTheOpenedFile_Click(object sender, RoutedEventArgs e)
		{
			MainGrid.Cursor = Cursors.Wait;

			excelController?.CloseApp();
			excelController = new MainExcelController();

			string chosenSheet = Sheets.SelectedItem as string;
			Sheets.ItemsSource = excelController.GetWorkbookSheetsNames(choosenFile);

			// imitate user's selecting if before user was selected sheet and that sheet is valid for new setting of program
			Sheets.SelectedIndex = -1;
			if (string.IsNullOrEmpty(chosenSheet) == false && Sheets.Items.Contains(chosenSheet))
			{
				Sheets.SelectedItem = chosenSheet;
			}
			else
			{
				Sheets.Text = "Select sheet";
				SendLater.IsEnabled = false;
				SendNow.IsEnabled = false;
			}

			MainGrid.Cursor = null;
		}
	}
}