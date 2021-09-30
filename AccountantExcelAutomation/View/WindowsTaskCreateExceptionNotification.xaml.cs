using System;
using System.Collections.Generic;
using System.Linq;
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

namespace AccountantExcelAutomation.View
{
	/// <summary>
	/// Логика взаимодействия для WindowsTaskCreateExceptionNotification.xaml
	/// </summary>
	public partial class WindowsTaskCreateExceptionNotification : Window
	{
		public WindowsTaskCreateExceptionNotification(Window owner, string programPath, string programParams, DateTime sendingDateTime)
		{
			this.Owner = owner;
			InitializeComponent();

			path.Text += programPath;
			_params.Text += programParams;
			dateTime.Text += sendingDateTime.ToString("G");
		}

		private void Window_Unloaded(object sender, RoutedEventArgs e)
		{
			this.DialogResult = true;
		}

		private void _MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
		{
			this.DragMove();
		}

		private void Button_Click(object sender, RoutedEventArgs e)
		{
			this.Close();
		}
	}
}
