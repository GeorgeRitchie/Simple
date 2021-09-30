using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using System.Linq;

using _base.Model;

namespace _base.Controller
{
	class ProgramConfigurationContext : IConfigurationContext
	{
		private readonly Dictionary<string, int> attachedFileExtentions = new Dictionary<string, int> { ["pdf"] = 57, ["xlsx"] = 51 }; // ["png"] = 0,
		private ProgramConfiguration configuration;
		private readonly string ConfigurationFilePath = GlobalObjects.Properties.ConfigurationFilesDirectoryName + "\\Configuration.json";

		public ProgramConfigurationContext()
		{
			if ((configuration = GetConfigurationFromFile()) == null)
			{
				configuration = new ProgramConfiguration();
				SetDefaultValues(configuration);
			}
		}

		private ProgramConfiguration GetConfigurationFromFile()
		{
			ProgramConfiguration configuration = null;

			using (StreamReader reader = new StreamReader(new FileStream(ConfigurationFilePath, FileMode.OpenOrCreate)))
			{
				string tempStr = reader.ReadToEnd();
				if (!string.IsNullOrEmpty(tempStr))
					configuration = JsonSerializer.Deserialize<ProgramConfiguration>(tempStr, new JsonSerializerOptions() { WriteIndented = true });
			}

			return configuration;
		}

		private void SetDefaultValues(ProgramConfiguration configuration)
		{
			configuration.MailTitle = "Salary";
			configuration.MailText = "Salary for current month. See attached file.";
			configuration.AttachedFileExtention = "pdf";
			configuration.AttachedFileName = "ReceiverDataFile";
			configuration.StartTrigger = "Start";
			configuration.EndTrigger = "End";
		}

		string IConfigurationContext.MailSenderEAddress
		{
			get
			{
				return configuration.MailSenderEAddress;
			}
			set
			{
				if (ReceiverManipulator.IsEAddress(value))
				{
					configuration.MailSenderEAddress = value;
				}
				else
				{
					throw new ArgumentException("Invalid mail format");
				}
			}
		}
		string IConfigurationContext.MailTitle { get => configuration.MailTitle; set => configuration.MailTitle = value; }
		string IConfigurationContext.MailText { get => configuration.MailText; set => configuration.MailText = value; }
		List<string> IConfigurationContext.AttachedFileExtentions => attachedFileExtentions.Keys.ToList();
		string IConfigurationContext.AttachedFileName { get => configuration.AttachedFileName; set => configuration.AttachedFileName = value; }
		int IConfigurationContext.ChosenAttachedFileExtentionAsInt => attachedFileExtentions[configuration.AttachedFileExtention];
		string IConfigurationContext.ChosenAttachedFileExtentionAsText
		{
			get
			{
				return configuration.AttachedFileExtention;
			}
			set
			{
				if (attachedFileExtentions.Keys.Contains(value))
				{
					configuration.AttachedFileExtention = value;
				}
				else
				{
					throw new ArgumentException("The passed parameter does not exists in the list of available extentions");
				}
			}
		}
		string IConfigurationContext.StartTrigger { get => configuration.StartTrigger; set => configuration.StartTrigger = value; }
		string IConfigurationContext.EndTrigger { get => configuration.EndTrigger; set => configuration.EndTrigger = value; }

		bool IConfigurationContext.SaveChanges()
		{
			try
			{
				WriteConfigurationsToFile(configuration);
			}
			catch
			{
				return false;
			}

			return true;
		}

		private void WriteConfigurationsToFile(ProgramConfiguration configuration)
		{
			if (configuration == null)
			{
				throw new ArgumentNullException();
			}

			string tempStr = JsonSerializer.Serialize(configuration, new JsonSerializerOptions() { WriteIndented = true });

			using (StreamWriter writer = new StreamWriter(new FileStream(ConfigurationFilePath, FileMode.Create)))
			{
				writer.Write(tempStr);
			}
		}
	}
}
