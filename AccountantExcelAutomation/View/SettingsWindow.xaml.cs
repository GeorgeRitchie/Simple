using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;
using _base.Controller;
using _base.Model;

namespace AccountantExcelAutomation.View
{
	/// <summary>
	/// Логика взаимодействия для SettingsWindow.xaml
	/// </summary>
	public partial class SettingsWindow : Window
	{
		public SettingsWindow(Window owner)
		{
			this.Owner = owner;
			InitializeComponent();

			if (GlobalObjects.Configuration?.MailSenderEAddress != null && GlobalObjects.Configuration?.MailSenderEAddress.Length > 0)
			{
				SetBinding(GlobalObjects.Configuration.MailSenderEAddress, "", EmailAddress, TextBox.TextProperty);
				isEmailCorrect = true;
			}

			SetBinding(GlobalObjects.Configuration.AttachedFileExtentions, "", AttachedFileExtension, ComboBox.ItemsSourceProperty);

			if (GlobalObjects.Configuration.ChosenAttachedFileExtentionAsText != null && GlobalObjects.Configuration.ChosenAttachedFileExtentionAsText.Length > 0)
			{
				AttachedFileExtension.SelectedItem = GlobalObjects.Configuration.ChosenAttachedFileExtentionAsText;
			}
			else
			{
				AttachedFileExtension.SelectedIndex = 0;
			}

			SetBinding(GlobalObjects.Configuration.MailTitle, "", MailTitle, TextBox.TextProperty);
			SetBinding(GlobalObjects.Configuration.AttachedFileName, "", AttachedFileName, TextBox.TextProperty);
			SetBinding(GlobalObjects.Configuration.MailText, "", MailText, TextBox.TextProperty);
			SetBinding(GlobalObjects.Configuration.StartTrigger, "", StartTrigger, TextBox.TextProperty);
			SetBinding(GlobalObjects.Configuration.EndTrigger, "", EndTrigger, TextBox.TextProperty);

			IReceiverContext receiverContext = IReceiverContext.CreateReceiverContext();
			ReceiversListData = receiverContext.GetReceivers().ToObservableCollection();
			SetBinding(ReceiversListData, "", ReceiversList, ListBox.ItemsSourceProperty);
		}

		private void SetBinding(object sourceObj, string sourceObjProperty, FrameworkElement recipient, DependencyProperty recipientProperty, BindingMode bindingMode = BindingMode.OneWay)
		{
			Binding binding = new Binding();

			binding.Source = sourceObj;
			binding.Path = new PropertyPath(sourceObjProperty);
			binding.Mode = bindingMode;
			recipient.SetBinding(recipientProperty, binding);
		}

		private void AttachedFileExtension_SelectionChanged(object sender, SelectionChangedEventArgs e)
		{
			try
			{
				GlobalObjects.Configuration.ChosenAttachedFileExtentionAsText = AttachedFileExtension.SelectedItem.ToString();
			}
			catch 
			{ }
		}

		private void MailTitle_LostFocus(object sender, RoutedEventArgs e)
		{
			GlobalObjects.Configuration.MailTitle = MailTitle.Text.Trim();
		}

		private void MailText_LostFocus(object sender, RoutedEventArgs e)
		{
			GlobalObjects.Configuration.MailText = MailText.Text.Trim();
		}

		private void StartTrigger_LostFocus(object sender, RoutedEventArgs e)
		{
			if (StartTrigger.Text != null && StartTrigger.Text.Length > 0)
			{
				GlobalObjects.Configuration.StartTrigger = StartTrigger.Text.Trim();
			}
			else
			{
				StartTrigger.Text = GlobalObjects.Configuration.StartTrigger;
			}
		}

		private void EndTrigger_LostFocus(object sender, RoutedEventArgs e)
		{
			if (EndTrigger.Text != null && EndTrigger.Text.Length > 0)
			{
				GlobalObjects.Configuration.EndTrigger = EndTrigger.Text.Trim();
			}
			else
			{
				EndTrigger.Text = GlobalObjects.Configuration.EndTrigger;
			}
		}

		private bool isEmailCorrect = false;
		private void EmailAddress_LostFocus(object sender, RoutedEventArgs e)
		{
			try
			{
				GlobalObjects.Configuration.MailSenderEAddress = EmailAddress.Text.Trim();
				isEmailCorrect = true;
			}
			catch
			{
				isEmailCorrect = false;
				EmailAddressPopup.IsOpen = true;
				EmailAddress.Text = EmailAddressValueBeforeChange;
			}
		}
		private void ClosePopup(object obj)
		{
			if (obj is Popup popup)
			{
				popup.IsOpen = false;
				popup.StaysOpen = false;
			}
		}

		private void EmailAddressPopup_Opened(object sender, EventArgs e)
		{
			if (EmailAddressPopup.IsOpen == true)
			{
				EmailAddressPopup.StaysOpen = true;

				TimerCallback closePopupCallback = new TimerCallback((obj) => this.Dispatcher.Invoke(() => ClosePopup(obj))); ;
				Timer timer = new Timer(closePopupCallback, EmailAddressPopup, 3000, Timeout.Infinite);
			}
		}

		private string EmailAddressValueBeforeChange;
		private void EmailAddress_GotFocus(object sender, RoutedEventArgs e)
		{
			EmailAddressValueBeforeChange = EmailAddress.Text;
			if (EmailAddress.Text == "Enter email address...")
			{
				EmailAddress.Text = "";
			}

			if (EmailAddressPopup.IsOpen == true)
			{
				EmailAddressPopup.IsOpen = false;
				EmailAddressPopup.StaysOpen = false;
			}
		}

		private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
		{
			// this line rizes a lost_Focus event on any element that was selected
			ReceiversList.Focus();

			if (isEmailCorrect == false && this.Owner != null)
			{
				if (MessageBox.Show("Email is not set correctly!\nDo you want to close settings window? (If there was email before, it will stay. Otherwise it may cause trouble in the future)", "ExcelAuto Warning", MessageBoxButton.YesNo, MessageBoxImage.Warning, MessageBoxResult.No, MessageBoxOptions.ServiceNotification) == MessageBoxResult.No)
				{
					e.Cancel = true;
					if (EmailAddressValueBeforeChange != "Enter email address...")
					{
						// if EmailAddressValueBeforeChange is not equal to default value it means there is correct email address in configuration and it will stay because it's imposible to save incorrect email address
						isEmailCorrect = true;
					}
				}
			}

			if (!GlobalObjects.Configuration.SaveChanges())
			{
				MessageBox.Show("Error was occurred while saving data! Data is not saved!", "ExcelAuto Error", MessageBoxButton.OK, MessageBoxImage.Error, MessageBoxResult.No, MessageBoxOptions.ServiceNotification);
			}
		}

		private ObservableCollection<Receiver> ReceiversListData = null;

		private void Add_Click(object sender, RoutedEventArgs e)
		{
			MainElementsGrid.Visibility = Visibility.Hidden;
			ReceiverEditorGrid.Visibility = Visibility.Visible;

			Operate_button.Content = "Add";

			isNameCorrectEntered = false;
			isEmailCorrectEntered = false;
		}

		private void AddReceiverHandler(string id, string name, string eAddress)
		{
			IReceiverContext receiverContext = IReceiverContext.CreateReceiverContext();
			if (name != null && name.Length > 0)
			{
				if (receiverContext.Create(name, eAddress))
				{
					// I did this because linq.Except gives stange results
					var newReceiversList = receiverContext.GetReceivers();
					var oldReceiversIDList = ReceiversListData.Select(i => i.ID).ToList();
					var difference = newReceiversList.Where(t => !oldReceiversIDList.Contains(t.ID)).FirstOrDefault();

					ReceiversListData.Add(difference);
				}
			}
		}

		private void Update_Click(object sender, RoutedEventArgs e)
		{
			if (ReceiversList.SelectedItem == null)
			{
				return;
			}

			MainElementsGrid.Visibility = Visibility.Hidden;
			ReceiverEditorGrid.Visibility = Visibility.Visible;

			ReceiverIDEditor.Text = ((Receiver)ReceiversList.SelectedItem).ID.ToString();
			ReceiverEmailAddress.Text = ((Receiver)ReceiversList.SelectedItem).EAddress.ToString();
			ReceiverNameEditor.Text = ((Receiver)ReceiversList.SelectedItem).Name.ToString();

			Operate_button.Content = "Update";

			isNameCorrectEntered = true;
			isEmailCorrectEntered = true;
		}

		private void UpdateReceiverHandler(string id, string name, string eAddress)
		{
			IReceiverContext receiverContext = IReceiverContext.CreateReceiverContext();
			if (receiverContext.Update(long.Parse(id), name, eAddress))
			{
				var tempReceiver = ReceiversListData.First(i => i.ID == long.Parse(id));
				tempReceiver.Name = name;
				tempReceiver.EAddress = eAddress;

				// this code is required because ReceiversListData's elements do not implement INotifyPropertyChanged
				int index = ReceiversListData.IndexOf(tempReceiver);
				ReceiversListData.RemoveAt(index);
				ReceiversListData.Insert(index, tempReceiver);
			}
		}

		private void Delete_Click(object sender, RoutedEventArgs e)
		{
			if (ReceiversList.SelectedItem == null)
				return;

			MessageBoxResult result = MessageBox.Show($"Do you want to delete \n{((Receiver)ReceiversList.SelectedItem).Name}", "Warning", MessageBoxButton.YesNo, MessageBoxImage.Question, MessageBoxResult.No, MessageBoxOptions.ServiceNotification);

			if (result == MessageBoxResult.Yes)
			{
				IReceiverContext receiverContext = IReceiverContext.CreateReceiverContext();
				if (receiverContext.Delete(((Receiver)ReceiversList.SelectedItem).ID))
				{
					ReceiversListData.Remove((Receiver)ReceiversList.SelectedItem);
				}
			}
		}

		bool isEmailCorrectEntered = false;
		bool isNameCorrectEntered = false;
		private void Operate_button_Click(object sender, RoutedEventArgs e)
		{
			if (isEmailCorrectEntered && isNameCorrectEntered)
			{
				if (Operate_button.Content.ToString() == "Update")
				{
					UpdateReceiverHandler(ReceiverIDEditor.Text, ReceiverNameEditor.Text, ReceiverEmailAddress.Text);
				}
				else if (Operate_button.Content.ToString() == "Add")
				{
					AddReceiverHandler(null, ReceiverNameEditor.Text, ReceiverEmailAddress.Text);
				}

				Cancel_Click(sender, e);
			}
			else
			{
				Warnings.Visibility = Visibility.Visible;
			}
		}

		private void Operate_button_LostFocus(object sender, RoutedEventArgs e)
		{
			Warnings.Visibility = Visibility.Hidden;
		}

		private void Cancel_Click(object sender, RoutedEventArgs e)
		{
			MainElementsGrid.Visibility = Visibility.Visible;
			ReceiverEditorGrid.Visibility = Visibility.Hidden;

			Operate_button.Content = "Operate";
			ReceiverIDEditor.Text = "-";
			ReceiverEmailAddress.Text = "Enter receiver email address...";
			ReceiverNameEditor.Text = "Enter receiver name...";
		}

		private void ReceiverNameEditor_LostFocus(object sender, RoutedEventArgs e)
		{
			if (ReceiverNameEditor.Text == "")
			{
				ReceiverNameEditor.Text = "Enter receiver name...";
				return;
			}

			if (char.IsLetter(ReceiverNameEditor.Text, 0) && ReceiverNameEditor.Text.Length > 0)
			{
				isNameCorrectEntered = true;
			}
			else
			{
				isNameCorrectEntered = false;
			}
		}

		private void ReceiverEmailAddress_LostFocus(object sender, RoutedEventArgs e)
		{
			if (ReceiverEmailAddress.Text == "")
			{
				ReceiverEmailAddress.Text = "Enter receiver email address...";
				return;
			}

			if (ReceiverManipulator.IsEAddress(ReceiverEmailAddress.Text))
			{
				isEmailCorrectEntered = true;
			}
			else
			{
				isEmailCorrectEntered = false;
			}
		}

		private void ReceiverNameEditor_GotFocus(object sender, RoutedEventArgs e)
		{
			if (ReceiverNameEditor.Text == "Enter receiver name...")
			{
				ReceiverNameEditor.Text = "";
			}
		}

		private void ReceiverEmailAddress_GotFocus(object sender, RoutedEventArgs e)
		{
			if (ReceiverEmailAddress.Text == "Enter receiver email address...")
			{
				ReceiverEmailAddress.Text = "";
			}
		}

		private void TextBox_KeyDown(object sender, KeyEventArgs e)
		{
			if (e.Key == Key.Enter)
			{
				e.Handled = true;
				ReceiversList.Focus();
			}
		}

		private void ReceiverTextBoxEditor_KeyDown(object sender, KeyEventArgs e)
		{
			if (e.Key == Key.Enter)
			{
				e.Handled = true;
				Operate_button.Focus();
			}
		}

		private void AttachedFileName_LostFocus(object sender, RoutedEventArgs e)
		{
			GlobalObjects.Configuration.AttachedFileName = AttachedFileName.Text.Trim();
		}
	}

	static class MethodsExtention
	{
		public static ObservableCollection<T> ToObservableCollection<T>(this List<T> receivers)
		{
			ObservableCollection<T> obReceivers = new ObservableCollection<T>();

			if (receivers == null || receivers.Count == 0)
				return obReceivers;

			foreach (var item in receivers)
			{
				obReceivers.Add(item);
			}

			return obReceivers;
		}
	}

}
