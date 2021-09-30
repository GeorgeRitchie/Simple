using Microsoft.Win32.TaskScheduler;
using System.Diagnostics;
using _base.Model;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;

namespace _base.Controller
{
	class WindowsTaskScheduler_Manager
	{
		private static readonly string TASK_FILE_FULL_NAME = GlobalObjects.Properties.HistoryOfSendingDirectoryName + "\\Tasks.json";

		public static void CreateTask(WindowsTaskScheduler_Object taskObj)
		{
			using (TaskService taskService = new TaskService())
			{
				string programToRun = Process.GetCurrentProcess().MainModule.FileName;
				string paramsToProgram = GlobalObjects.Properties.ConsoleProgramKey_Send + " " + taskObj.TaskID;

				TaskDefinition taskDefinition = taskService.NewTask();
				taskDefinition.RegistrationInfo.Description = "Send Deferred mails";

				taskDefinition.Triggers.Add(new TimeTrigger(taskObj.DateTimeOfSendingMail.ToUniversalTime()));

				taskDefinition.Actions.Add(new ExecAction(programToRun, paramsToProgram, null));

				string taskLocationAndNameInWindowsTaskScheduler = $"\\SendMailTasks\\{taskObj.DateTimeOfSendingMail.ToShortDateString()} - {taskObj.TaskID}";
				taskService.RootFolder.RegisterTaskDefinition(taskLocationAndNameInWindowsTaskScheduler, taskDefinition);
			}
		}

		public static List<WindowsTaskScheduler_Object> GetExistingTasks()
		{
			List<WindowsTaskScheduler_Object> taskObjects = null;

			using (StreamReader strRead = new StreamReader(new FileStream(TASK_FILE_FULL_NAME, FileMode.OpenOrCreate)))
			{
				string str = strRead.ReadToEnd();

				if (str != null && str.Length > 0)
					taskObjects = JsonSerializer.Deserialize<List<WindowsTaskScheduler_Object>>(str, new JsonSerializerOptions() { WriteIndented = true });
				else
					taskObjects = new List<WindowsTaskScheduler_Object>();
			}

			return taskObjects;
		}

		public static void UpdateFileOfWindowsTasks(List<WindowsTaskScheduler_Object> listOfTaskObjects)
		{
			using (StreamWriter strWrite = new StreamWriter(new FileStream(TASK_FILE_FULL_NAME, FileMode.Create)))
			{
				strWrite.Write(JsonSerializer.Serialize(listOfTaskObjects, new JsonSerializerOptions() { WriteIndented = true }));
			}
		}
	}
}
