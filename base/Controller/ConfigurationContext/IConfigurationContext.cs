using System.Collections.Generic;

namespace _base.Controller
{
	public interface IConfigurationContext
	{
		bool SaveChanges();

		public string MailSenderEAddress { get; set; }
		public string MailTitle { get; set; }
		public string MailText { get; set; }
		public List<string> AttachedFileExtentions { get; }
		public int ChosenAttachedFileExtentionAsInt { get; }
		public string ChosenAttachedFileExtentionAsText { get; set; }
		public string AttachedFileName { get; set; }
		public string StartTrigger { get; set; }
		public string EndTrigger { get; set; }
	}
}
