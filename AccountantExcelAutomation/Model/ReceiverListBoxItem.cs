using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;

namespace AccountantExcelAutomation.Model
{

	class ReceiverListBoxItem : INotifyPropertyChanged
	{
		private string name = "";
		private bool? isChecked = false;

		public event PropertyChangedEventHandler PropertyChanged;
		public void OnPropertyChanged([CallerMemberName] string prop = "")
		{
			PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(prop));
		}

		public string Name
		{
			get
			{
				return name;
			}
			set
			{
				name = value;
				OnPropertyChanged("Name");
			}
		}
		public bool? IsChecked
		{
			get
			{
				return isChecked;
			}
			set
			{
				isChecked = value;
				OnPropertyChanged("IsChecked");
			}
		}

		public static int CheckedCount { get; set; }
	}
}
