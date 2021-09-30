using System.IO;

namespace _base.Controller
{
	class DirectoryManager
	{
		public DirectoryInfo GetDirectoryAndCreateItIfDoesnotExists(string directoryFullName)
		{
			DirectoryInfo givenDirectory = new DirectoryInfo(directoryFullName);
			if (!givenDirectory.Exists)
			{
				givenDirectory.Create();
			}

			return givenDirectory;
		}
	}
}
