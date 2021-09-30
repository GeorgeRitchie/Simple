namespace _base.Model
{
	public class ProgramProperties
	{
		public bool ExcelAppVisibility { get; set; } = false;
		
		public string ThisProgramDataFilesDirectory = string.Empty;
		public string HistoryOfSendingDirectoryName = string.Empty;
		public string ConfigurationFilesDirectoryName = string.Empty;
		public string LogFilesDirectoryName = string.Empty;
		
		public string ConsoleProgramKey_Send = "-send";
		public string ConsoleProgramKey_Configure = "-configure";
	}
}
