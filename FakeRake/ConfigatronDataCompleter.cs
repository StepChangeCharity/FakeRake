
using System;
using System.Text.RegularExpressions;
using System.Collections.Generic;


namespace FakeRake
{
	/// <summary>
	/// Description of ConfigatronDataCompleter.
	/// </summary>
	public class ConfigatronDataCompleter : Interfaces.IConfigDataCompleter
	{
		public Dictionary<string, string> SettingsDictionary {get; set;}
		public string CompleteData(string sourceData) 
		{
			var configData = Regex.Replace(sourceData, @"#\{(.*?)\}", HandleReplacement);
			return configData;
		}
		
		protected string HandleReplacement(Match m)
		{
			var key = m.Groups[1].Value.ToUpper();
			var match = Regex.Match(key,"^configatron\\.(.*)$", RegexOptions.IgnoreCase);
			if (match.Success){
				key = match.Groups[1].Value;
			}
			var result = m.Value;
			if (this.SettingsDictionary.ContainsKey(key))
			{
				result = this.SettingsDictionary[key];
			}
			return result;
		}
		
	}
}
