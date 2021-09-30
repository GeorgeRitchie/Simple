namespace _base.Controller
{
	class MailItem
	{
		public string Subject { get; set; }
		public string HtmlBody { get; set; }
		public string Recipients { get; set; }
		public string[] FilePaths { get; set; }
		public string SenderEAddress { get; set; }
	}
}
