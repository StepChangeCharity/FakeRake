
using System;
using System.Linq;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Threading.Tasks;
using System.IO;
using System.Text.RegularExpressions;

namespace FakeRake
{
	
	
	class Program
	{
		
		protected static Dictionary<string, string> _settingsDictionary;
		protected static List<Task> _tasks = new List<Task>();
		
		public static void Main(string[] args)
		{
			
			// Get the environment name...
			var environmentName = GetEnvironmentName();
			
			// Now build the dictionary...
			_settingsDictionary = GetSettingsDictionary(environmentName);
			
			// Now run from this folder downward
			ProcessFolder(System.IO.Directory.GetCurrentDirectory());
			
			
			
			// Wait for everything to complete...
			Task.WaitAll(_tasks.ToArray());
			
			
			// TODO: Implement Functionality Here
			
			Console.Write("Press any key to continue . . . ");
			Console.ReadKey(true);
		}
		
		private static string GetEnvironmentName()
		{
			var args = System.Environment.GetCommandLineArgs();
			if (args.Count() != 2)
			{
				throw new ArgumentException("You have failed to supply the correct number of command line parameters (i.e. 1 - the name of the environment to use)");
			}
			return args[1];
		}
		
		private static  Dictionary<string, string> GetSettingsDictionary(string environmentName)
		{
			var configatronPath = GetConfigatronPath();
			var result = new Dictionary<string, string>();
			string connectionString = string.Format("Provider=Microsoft.Jet.OLEDB.4.0;Data Source={0};Extended Properties=\"Excel 8.0;HDR=YES;\"", configatronPath);
			using (var conn = new OleDbConnection(connectionString))
			{
				conn.Open();
				using (var cmd = new OleDbCommand())
				{
					cmd.CommandText = "Select * from [Sheet1$]";
					cmd.CommandType = CommandType.Text;
					cmd.Connection = conn;
					var reader = cmd.ExecuteReader();
					var configFieldIndex = 0;
					int dataFieldIndex;
					try{
						dataFieldIndex = reader.GetOrdinal(environmentName.ToUpper());
					}
					catch (IndexOutOfRangeException)
					{
						throw new ApplicationException(string.Format("Unable to find environment {0} in the configuration.xls file...", environmentName));
					}
					
					
					while (reader.Read())
					{
						var key = reader.GetString(configFieldIndex).ToUpper();
						var value = reader.GetValue(dataFieldIndex).ToString();
						result[key] = value;
					}
					
				}
			}
			
			return result;
		}
		
		private static string GetConfigatronPath() 
		{
			var result =  string.Format("{0}\\Configatron.xls",System.IO.Directory.GetCurrentDirectory());
			if (!System.IO.File.Exists(result)){
				throw new ApplicationException(string.Format("Unable to find the configatron excel spreadsheet in this location - {0}", result));
			}
			
			return result;
			
			
		}
		
		protected static void ProcessFolder(string folderPath){
			var folder = new System.IO.DirectoryInfo(folderPath);
			var configatronFiles = folder.GetFiles("*.configatron");
			
			foreach (var file in configatronFiles)
			{
				_tasks.Add(GetConfigurationFileTask(file.FullName));
			}
			
			
			// And now for the recursion bit...
			foreach(var subFolder in folder.GetDirectories().Select(x => x.FullName))
			{
				ProcessFolder(subFolder);
			}
		}
		
		
		protected  static Task GetConfigurationFileTask(string filename)
		{
			var task = new Task(() => ProcessConfigatronFile(filename));
			task.Start();
			return task;
			
		}
		
		protected static void ProcessConfigatronFile(string filename)
		{
			
			string configPath =System.Text.RegularExpressions.Regex.Match(filename, "(.*)\\.configatron",System.Text.RegularExpressions.RegexOptions.IgnoreCase).Groups[1].Value
				+ ".config";
			
			string data = File.ReadAllText(filename);
			var completer = new ConfigatronDataCompleter() {SettingsDictionary = _settingsDictionary};
			var configData = completer.CompleteData(data);
			
			if(File.Exists(configPath))
			{
				var attributes = File.GetAttributes(configPath);
				attributes = attributes & (~FileAttributes.ReadOnly);
				File.SetAttributes(configPath, attributes);
			}
			
			File.WriteAllText(configPath, configData);
			
			return;
		}
		
		
	}
}