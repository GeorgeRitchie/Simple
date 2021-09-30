using System;
using System.Text;
using System.IO;
using System.Text.Json;

using _base.Model;

namespace _base.Controller
{
	public static class Logger
	{
		class LoggerObject
		{
			public DateTime timeOfLogging { get; set; }
			public string objCalledLogger { get; set; }
			public string loggingMsg { get; set; }
		}

		enum BeforeAfter
		{
			Before,
			After
		}

		private const string LOG_FILE_NAME = "log.json";
		private static FileStream file = null;

		public static void Log(string msg, string sender)
		{
			file = new FileStream(GlobalObjects.Properties.LogFilesDirectoryName + "\\" + LOG_FILE_NAME, FileMode.OpenOrCreate, FileAccess.ReadWrite);

			LoggerObject loggerObject = new LoggerObject() { loggingMsg = msg, objCalledLogger = sender, timeOfLogging = DateTime.Now };

			if (IsLogFileEmpty())
			{
				WriteText(JsonSerializer.Serialize(loggerObject, new JsonSerializerOptions() { WriteIndented = true }), true);
			}
			else
			{
				SetSeekPositionToWrite();
				WriteText(JsonSerializer.Serialize(loggerObject, new JsonSerializerOptions() { WriteIndented = true }), false);
			}

			file.Close();
		}

		private static bool IsLogFileEmpty()
		{
			return file.Length == 0;
		}

		private static void WriteText(string str, bool isItFirstLogInFile)
		{
			string beginOfSRT = "";

			if (isItFirstLogInFile)
			{
				beginOfSRT = "[\n";
			}
			else
			{
				beginOfSRT = ",\n";
			}

			string endOfSRT = "\n]\n";

			file.Write(Encoding.Default.GetBytes(beginOfSRT + str + endOfSRT));
			file.Flush();
		}

		private static void SetSeekPositionToWrite()
		{
			SetSeekPositionToEnd();
			SetSeekPositionBeforeLastSquareBrackets();
			SetSeekPositionAfterLastBraces();
		}

		private static void SetSeekPositionToEnd()
		{
			file.Seek(-1, SeekOrigin.End);
		}

		private static void SetSeekPositionBeforeLastSquareBrackets()
		{
			SetSeekPositionOverSymbol(']', BeforeAfter.Before);
		}

		private static void SetSeekPositionAfterLastBraces()
		{
			SetSeekPositionOverSymbol('}', BeforeAfter.After);
		}

		private static void SetSeekPositionOverSymbol(char symbol, BeforeAfter beforeAfter)
		{
			byte[] bit = new byte[1];

			while (true)
			{
				file.Read(bit, 0, 1);

				if (Convert.ToChar(bit[0]) == symbol)
				{
					if (beforeAfter == BeforeAfter.Before)
					{
						file.Seek(-1, SeekOrigin.Current);
					}

					break;
				}

				file.Seek(-2, SeekOrigin.Current);
			}
		}
	}
}
