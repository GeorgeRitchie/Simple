using System.Text.Json.Serialization;

using Excel = Microsoft.Office.Interop.Excel;

namespace _base.Model
{
	class Mail_Object
	{
		public long ReceiverID { get; set; }
		public string ReceiverName { get; set; }
		public string ReceiverEAddress { get; set; }
		public string MailTitle { get; set; }
		public string MailText { get; set; }
		public string AttachedFile_FullName { get; set; }

		[JsonIgnore]
		public Excel.Range DataRangeOfReceiver { get; set; }

		[JsonPropertyName("DataRange")]
		// for reasons I don't know JsonSerializer.Serialize cannot serialize Excel.Range,
		// so I need only ranges address as string so save in file,
		// for being able to indentify this receiver's data in Excel File
		public string RangesDataForJson
		{
			get
			{
				if (DataRangeOfReceiver != null)
					return dataRangeOfReceiverForJson = DataRangeOfReceiver.Address.ToString();
				else
					return dataRangeOfReceiverForJson;
			}
			set
			{
				dataRangeOfReceiverForJson = value;
			}
		}
		private string dataRangeOfReceiverForJson;
	}
}
