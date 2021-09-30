using System;
using System.Text.Json;

using System.IO;
using _base.Model;

namespace _base.Controller
{
	class MailDataRecorder
	{
		const string SENDING_MAILS_DATA_FILE_NAME = "Data.json";
		string ATTACHED_FILE_NAME = GlobalObjects.Configuration.AttachedFileName;
		const string RECEIVERS_FILES_FOLDER_NAME = "Receivers";
		private string SENDING_MAILS_DATA_FILE_DIRECTORY_NAME = GlobalObjects.Properties.HistoryOfSendingDirectoryName;
		DataOfSendingMails_OfSelectedSheet dataOfSendingMails;

		public MailDataRecorder(DataOfSendingMails_OfSelectedSheet dataOfSendingMails)
		{
			this.dataOfSendingMails = dataOfSendingMails;
		}

		public void CreateDirectory()
		{
			if (dataOfSendingMails.Mails != null)
			{
				DirectoryInfo directoryOfAllReceiversData = CreateOrRecreateGeneralDirectories();
				CreateDirectoriesForEachReceiverAndSaveReceiverFileLocation(directoryOfAllReceiversData);
			}
			else
			{
				Logger.Log("There is no data of receivers to create directory for mails and data of each receiver.", typeof(MailDataRecorder).FullName + '.' + nameof(CreateDirectory));
			}
		}

		private DirectoryInfo CreateOrRecreateGeneralDirectories()
		{
			SENDING_MAILS_DATA_FILE_DIRECTORY_NAME += "\\" + dataOfSendingMails.SelectedSheetName + " " + dataOfSendingMails.ExcelFileName;

			DirectoryManager directoryManager = new DirectoryManager();

			DirectoryInfo sendingMailsDataFileDirectory = directoryManager.GetDirectoryAndCreateItIfDoesnotExists(SENDING_MAILS_DATA_FILE_DIRECTORY_NAME);
			DeleteDirectorysAllContaintmentsIfExits(sendingMailsDataFileDirectory);

			// saving directory to write data later in method RecordAllData
			dataOfSendingMails.DataOfSendingMailsFileFullName = SENDING_MAILS_DATA_FILE_DIRECTORY_NAME + "\\" + SENDING_MAILS_DATA_FILE_NAME;

			DirectoryInfo directoryOfAllReceiversData = directoryManager.GetDirectoryAndCreateItIfDoesnotExists(sendingMailsDataFileDirectory.FullName + "\\" + RECEIVERS_FILES_FOLDER_NAME);

			return directoryOfAllReceiversData;
		}

		private void DeleteDirectorysAllContaintmentsIfExits(DirectoryInfo directory)
		{
			if (directory.Exists && (directory.GetFiles().Length > 0 || directory.GetDirectories().Length > 0))
			{
				string directoryFullName = directory.FullName;
				directory.Delete(true);
				directory = new DirectoryInfo(directoryFullName);
			}
		}

		// TODO в книге "чистый код" сказано чтобы метод не делал побочных действий, или надо указать побочное действие в имени
		//		или убрать действие в другое место
		//		я не смог найти как убрать побочное действие без дублирования цикла, поэтому включил в имя упоминание об побочном действии
		//		КАКОВО ВАШЕ МНЕНИЕ?
		private void CreateDirectoriesForEachReceiverAndSaveReceiverFileLocation(DirectoryInfo directoryOfAllReceiversData)
		{
			DirectoryManager directoryManager = new DirectoryManager();
			DirectoryInfo directoryOfCurrentReceiver = null;

			foreach (Mail_Object receiver in dataOfSendingMails.Mails)
			{
				directoryOfCurrentReceiver = directoryManager.GetDirectoryAndCreateItIfDoesnotExists(directoryOfAllReceiversData.FullName + "\\" + receiver.ReceiverName);

				SaveCurrentReceiverAttachedFileFullName(receiver, directoryOfCurrentReceiver);
			}
		}

		private void SaveCurrentReceiverAttachedFileFullName(Mail_Object receiver, DirectoryInfo receiverDataFileDirectory)
		{
			receiver.AttachedFile_FullName = receiverDataFileDirectory.FullName + "\\" + ATTACHED_FILE_NAME + "." + GlobalObjects.Configuration.ChosenAttachedFileExtentionAsText;
		}



		public bool RecordAllData()
		{
			bool result = true;

			if (!IsDirectoryCreated())
			{
				Logger.Log("There is no directory created to record mails and data of each receiver.", typeof(MailDataRecorder).FullName + '.' + nameof(RecordAllData));
				return false;
			}

			try
			{
				WriteToFile();
			}
			catch (Exception e)
			{
				result = false;
				Logger.Log("The exception was thrown during recording all data of consignment of mails and receivers. Recording was not successful! Exception message: " + e.Message, typeof(MailDataRecorder).FullName + '.' + nameof(RecordAllData));
			}

			return result;
		}

		private bool IsDirectoryCreated()
		{
			FileInfo fileWithDataOfSendingMails = new FileInfo(dataOfSendingMails.DataOfSendingMailsFileFullName);
			return fileWithDataOfSendingMails.Directory.Exists;
		}

		private void WriteToFile()
		{
			FileInfo fileWithDataOfSendingMails = new FileInfo(dataOfSendingMails.DataOfSendingMailsFileFullName);

			using (StreamWriter recordDataObj = new StreamWriter(fileWithDataOfSendingMails.Open(FileMode.OpenOrCreate)))
			{
				recordDataObj.Write(JsonSerializer.Serialize(dataOfSendingMails, new JsonSerializerOptions() { WriteIndented = true }));
			}
		}
	}
}
