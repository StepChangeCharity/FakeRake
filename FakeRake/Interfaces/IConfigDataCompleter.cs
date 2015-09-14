
using System;
using System.Collections.Generic;
namespace FakeRake.Interfaces
{
	/// <summary>
	/// Description of IConfigDataCompleter.
	/// </summary>
	public interface IConfigDataCompleter
	{
		Dictionary<string, string> SettingsDictionary {get; set;}
		string CompleteData(string sourceData);
		
	}
}
