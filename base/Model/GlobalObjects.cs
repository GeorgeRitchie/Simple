using System;
using _base.Controller;

namespace _base.Model
{
	public static class GlobalObjects
	{
		private const string ProgramDataFilesDirectoryName = "ExcelAuto";
		private const string ProgramHistoryFilesDirectoryName = "History";
		private const string ProgramConfigurationFilesDirectoryName = "Configuration";
		private const string ProgramLogFilesDirectoryName = "Log";

		public static IConfigurationContext Configuration { get; set; }
		public static ProgramProperties Properties { get; set; }

		static GlobalObjects()
		{
			Properties = new ProgramProperties();
			Properties.ThisProgramDataFilesDirectory = GetPathOfProgramDataFilesDirectoryName();
			Properties.HistoryOfSendingDirectoryName = GetPathOfGivenDirectoryName(ProgramHistoryFilesDirectoryName);
			Properties.ConfigurationFilesDirectoryName = GetPathOfGivenDirectoryName(ProgramConfigurationFilesDirectoryName);
			Properties.LogFilesDirectoryName = GetPathOfGivenDirectoryName(ProgramLogFilesDirectoryName);
			Configuration = new ProgramConfigurationContext();
		}

		private static string GetPathOfProgramDataFilesDirectoryName()
		{
			DirectoryManager directoryManager = new DirectoryManager();
			return directoryManager.GetDirectoryAndCreateItIfDoesnotExists(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "\\" + ProgramDataFilesDirectoryName).FullName;
		}

		private static string GetPathOfGivenDirectoryName(string directoryName)
		{
			DirectoryManager directoryManager = new DirectoryManager();
			return directoryManager.GetDirectoryAndCreateItIfDoesnotExists(Properties.ThisProgramDataFilesDirectory + "\\" + directoryName).FullName;
		}
	}
}
