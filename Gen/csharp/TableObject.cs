using System;
using Newtonsoft.Json;

namespace Jrainstar.Table
{
	public class TableObject<T>
	{
		private static Dictionary<int, T> curTable;

		private static Dictionary<string, Dictionary<int, T>> tagTable;

		public static T Get(int id, string tag = "")
		{
			if (string.IsNullOrEmpty(tag))
			{
				if (curTable == null)
				{
					curTable = JsonConvert.DeserializeObject<Dictionary<int, T>>(TableBuilder.Load(typeof(T).Name));
				}
				curTable.TryGetValue(id, out T value);
				return value;
			}
			else
			{
				if (tagTable == null)
				{
					tagTable = new Dictionary<string, Dictionary<int, T>>();
				}
					if (!tagTable.ContainsKey(tag))
					{
						tagTable[tag] = JsonConvert.DeserializeObject<Dictionary<int, T>>(TableBuilder.Load($"{typeof(T).Name}_{tag}"));
					}
					tagTable[tag].TryGetValue(id, out T value);
					return value;
				}
			}
		}
	}
