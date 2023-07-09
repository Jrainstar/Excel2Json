using System;

namespace Jrainstar.Table
{
	public static class TableBuilder
	{
		public static Func<string, string> onLoad { get; set; }
		public static string Load(string fileName)
		{
			return onLoad?.Invoke(fileName);
		}
	}
}
