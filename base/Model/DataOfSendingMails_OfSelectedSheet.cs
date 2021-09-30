using System;
using System.Collections.Generic;
using System.Text.Json.Serialization;

using _base.Controller;

namespace _base.Model
{
	class DataOfSendingMails_OfSelectedSheet
	{
		public string ExcelFileName { get; set; }
		public string SelectedSheetName { get; set; }
		public DateTime MailsSendDateTime { get; set; }
		public bool AreMailsSent { get; set; } = false;
		public string DataOfSendingMailsFileFullName { get; set; }
		public List<Mail_Object> Mails { get; set; }

		[JsonIgnore]
		// Each mailDataRecorder saves directory for each Data_Object,
		// and to prevent using directory of one dataObject for another I did it as a property for dataObject
		internal MailDataRecorder MailDataRecorder { get; private set; }

		public DataOfSendingMails_OfSelectedSheet()
		{
			MailDataRecorder = new MailDataRecorder(this);
		}
	}
}
