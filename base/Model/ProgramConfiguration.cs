namespace _base.Model 
{
	class ProgramConfiguration
	{
		public string MailSenderEAddress { get; set; }
		public string MailTitle { get; set; }
		public string MailText { get; set; }
		public string AttachedFileExtention { get; set; }
		public string AttachedFileName { get; set; }
		public string StartTrigger { get; set; }
		public string EndTrigger { get; set; }
	}
}
