using System;
using System.Diagnostics;
using _base.Model;

namespace Configure
{
	class Program
	{
		static void Main(string[] args)
		{
			Process.Start($"AccountantExcelAutomation.exe", GlobalObjects.Properties.ConsoleProgramKey_Configure);
		}
	}
}
