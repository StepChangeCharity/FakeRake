﻿
using System;
using System.Linq;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Threading.Tasks;
using System.IO;
using System.Text.RegularExpressions;
using Excel = Microsoft.Office.Interop.Excel;

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
			string connectionString = string.Format("Provider=Microsoft.Jet.OLEDB.4.0;Data Source={0};Extended Properties=\"Excel 8.0;HDR=No;IMEX=1;\";", configatronPath);
			using (var conn = new OleDbConnection(connectionString))
			{
				conn.Open();
				
				string sql = "Select * from [Sheet1$]";
				using (var adaptor = new OleDbDataAdapter(sql,conn))
				{
					DataTable dt = new DataTable();
					adaptor.Fill(dt);
					var configFieldIndex = 0;
					int dataFieldIndex = -1;
					
					if (dt.Rows.Count > 0)
					{
						var dr = dt.Rows[0];
						
						for(int tempIndex = 0; tempIndex < dt.Columns.Count; tempIndex++)
						{
							var colName = dr[tempIndex].ToString();
							if (string.Equals(colName, environmentName.ToUpper(), StringComparison.InvariantCultureIgnoreCase)){
								dataFieldIndex = tempIndex;
								break;
							}
						}
					}
					
					if (dataFieldIndex == -1) throw new ApplicationException(string.Format("Unable to find environment {0} in the configuration.xls file...", environmentName));
					
					
					
					for (int tempIndex = 1; tempIndex < dt.Rows.Count; tempIndex++)
					{
						var key = dt.Rows[tempIndex][configFieldIndex].ToString().ToUpper();
						
						var value = dt.Rows[tempIndex][dataFieldIndex].ToString();
						result[key] = value;
						if (string.IsNullOrWhiteSpace(value)) 
						{
							result[key]="Whatever...";
						}
						
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